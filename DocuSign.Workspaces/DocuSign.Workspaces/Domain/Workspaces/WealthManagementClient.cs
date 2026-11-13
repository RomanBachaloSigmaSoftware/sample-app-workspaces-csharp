using DocumentFormat.OpenXml.EMMA;
using Docusign.IAM.SDK.Models.Components;
using DocuSign.eSign.Api;
using DocuSign.eSign.Client;
using DocuSign.eSign.Model;
using DocuSign.Workspaces.Domain.Workspaces.Models;
using DocuSign.Workspaces.Infrastructure.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Document = DocuSign.Workspaces.Domain.Workspaces.Models.Document;

namespace DocuSign.Workspaces.Domain.Workspaces;

public class WealthManagementClient(
    IDocuSignApiProvider docuSignApiProvider,
    IAppConfiguration appConfiguration,
    IAccountRepository accountRepository)
    : IWealthManagementClient
{
    public async Task<string> CreateWorkspaces(CreateWorkspacesModel createWorkspacesModel)
    {
        var workspaceBody = new CreateWorkspaceBody
        {
            Name = createWorkspacesModel.WorkspacesName,
        };
        var workspace = await docuSignApiProvider.Workspace2.CreateWorkspaceAsync(accountRepository.AccountId, workspaceBody);

        return workspace.WorkspaceId;
    }

    public async Task<List<EnvelopeModel>> AddSelectedDocumentsForClientPackage(WorkspaceAddDocumentsModel createModel)
    {
        var envelopes = new List<EnvelopeModel>();
        foreach (var document in createModel.Documents)
        {
            var documentId = await UploadWorkspaceDocument(document, createModel.WorkspaceId);
            var envelopeId = await ProcessDocument(createModel, document.Name, documentId);

            var envelopeModel = await UpdateEachEnvelopWithDocument(
                envelopeId: envelopeId,
                document: document,
                model: createModel);

            envelopes.Add(envelopeModel);
        }

        return envelopes;
    }

    private async Task<string> UploadWorkspaceDocument(Document document, string workspaceId)
    {
        var documentRequest = new AddWorkspaceDocumentRequest
        {
            File = new AddWorkspaceDocumentRequestFile
            {
                Content = Convert.FromBase64String(document.Base64String),
                FileName = document.Name
            }
        };
        var addedDocument = await docuSignApiProvider.WorkspaceDocuments.AddWorkspaceDocumentAsync(accountRepository.AccountId, workspaceId, documentRequest);

        return addedDocument.DocumentId;
    }

    private async Task<string> ProcessDocument(WorkspaceAddDocumentsModel createModel, string envelopeName, string documentId)
    {
        var workspace = await docuSignApiProvider.Workspace2.GetWorkspaceAsync(accountRepository.AccountId, createModel.WorkspaceId);

        var workspaceEnvelopeForCreate = new WorkspaceEnvelopeForCreate
        {
            EnvelopeName = envelopeName,
            DocumentIds = [documentId]
        };

        var envelopeResponse = await docuSignApiProvider.Workspace2.CreateWorkspaceEnvelopeAsync(
            accountRepository.AccountId,
            createModel.WorkspaceId,
            workspaceEnvelopeForCreate);

        var uploadRequestBody = new CreateWorkspaceUploadRequestBody
        {
            Name = envelopeName + " Upload " + Guid.NewGuid(),
            Description = "Description Test",
            DueDate = DateTime.Now.AddDays(7),
            Status = WorkspaceUploadRequestStatus.Draft,
            Assignments =
            [
                new CreateWorkspaceUploadRequestAssignment
                {
                    Email = createModel.PrimaryOwnerEmail,
                    FirstName = createModel.PrimaryOwnerFirstName,
                    LastName = createModel.PrimaryOwnerLastName,
                    UploadRequestResponsibilityTypeId = WorkspaceUploadRequestResponsibilityType.Assignee,
                    AssigneeUserId = workspace.CreatedByUserId
                }
            ]
        };
        await docuSignApiProvider.WorkspaceUploadRequest.CreateWorkspaceUploadRequestAsync(
            accountRepository.AccountId,
            createModel.WorkspaceId,
            uploadRequestBody);
        
        return envelopeResponse.EnvelopeId;
    }

    private async Task<EnvelopeModel> UpdateEachEnvelopWithDocument(string envelopeId, Document document, WorkspaceAddDocumentsModel model)
    {
        var accountsApi = new AccountsApi(docuSignApiProvider.ApiClient);
        var response = await accountsApi.GetAccountIdentityVerificationAsync(accountRepository.AccountId);

        var workflow = response.IdentityVerification.FirstOrDefault(x => x.DefaultName == "DocuSign ID Verification");
        if (workflow == null)
        {
            throw new ApiException(0, "IDENTITY_WORKFLOW_INVALID_ID");
        }

        var workflowId = workflow.WorkflowId;
        var eventNotificationUrl = $"{appConfiguration.DocuSign.EventNotificationBaseUrl}/api/callback/event";

        var signHere1 = new SignHere
        {
            AnchorString = "/sn1/",
            AnchorUnits = "pixels",
            AnchorXOffset = "10",
            AnchorYOffset = "20",
        };

        var signer1Tabs = new Tabs
        {
            SignHereTabs = [signHere1],
        };

        var recipientIdentityVerification = new RecipientIdentityVerification
        {
            WorkflowId = workflowId
        };

        var recipients = new Recipients
        {
            Signers =
            [
                new Signer
                {
                    Name = model.PrimaryOwnerFirstName + " " + model.PrimaryOwnerLastName,
                    Email = model.PrimaryOwnerEmail,
                    RoutingOrder = "1",
                    Status = "Created",
                    DeliveryMethod = "Email",
                    RecipientId = "1",
                    Tabs = signer1Tabs,
                    IdentityVerification = recipientIdentityVerification,
                }
            ]
        };

        if (!string.IsNullOrEmpty(model.SecondaryOwnerEmail))
        {
            recipients.Signers.Add(new Signer
            {
                Name = model.PrimaryOwnerFirstName + " " + model.PrimaryOwnerLastName,
                Email = model.PrimaryOwnerEmail,
                RoutingOrder = "1",
                Status = "Created",
                DeliveryMethod = "Email",
                RecipientId = "2",
                Tabs = signer1Tabs,
                IdentityVerification = recipientIdentityVerification,
            });
        }

        var envelope = new Envelope
        {
            EnvelopeIdStamping = "true",
            EmailSubject = "Please Sign",
            EmailBlurb = "Sample text for email body",
            EnvelopeId = envelopeId,
            Recipients = recipients,
            Status = "Sent"
        };

        await docuSignApiProvider.EnvelopApi.UpdateAsync(accountRepository.AccountId, envelopeId, envelope);

        return new EnvelopeModel(document.Name, envelope.Status);
    }
}
