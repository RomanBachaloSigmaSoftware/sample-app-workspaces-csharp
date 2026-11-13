using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocuSign.eSign.Model;
using Docusign.IAM.SDK.Models.Components;
using DocuSign.Workspaces.Domain.CarePlans.Model;
using DocuSign.Workspaces.Infrastructure.Services.Interfaces;

namespace DocuSign.Workspaces.Domain.CarePlans;

public class CarePlansService(IDocuSignApiProvider docuSignApiProvider, IAppConfiguration appConfiguration, IAccountRepository accountRepository)
    : ICarePlansService
{
    public async Task<List<PhysicianModel>> GetPhysician()
    {
        var physiciansWorkspaces = new List<PhysicianModel>();
        if (appConfiguration.DocuSign.TestAccountConnectionSettings.AccountId == accountRepository.AccountId)
        {
            List<string> physicians = ["Dr. Max Payne", "Dr. Angela Kerr", "Dr. Luke Heer"];
            var workspaces = await docuSignApiProvider.Workspace2.GetWorkspacesAsync(accountRepository.AccountId);
            if (workspaces.Workspaces != null || workspaces.Workspaces?.Count != 0)
            {
                var physician = workspaces.Workspaces
                    ?.Where(a => physicians.Contains(a.Name + " Workspace"))
                    .Select(a => new PhysicianModel
                    {
                        Name = a.Name,
                        WorkspaceId = a.WorkspaceId
                    }).ToList();

                if (physician?.Count == physicians.Count)
                {
                    return physician;
                }
            }

            foreach (var physician in physicians)
            {
                var workspaceBody = new CreateWorkspaceBody
                {
                    Name = physician + " Workspace"
                };
                var workspace = await docuSignApiProvider.Workspace2.CreateWorkspaceAsync(accountRepository.AccountId, workspaceBody);

                physiciansWorkspaces.Add(new PhysicianModel
                {
                    Name = physician,
                    WorkspaceId = workspace.WorkspaceId
                });
            }
        }
        else
        {
            var workspaceBody = new CreateWorkspaceBody
            {
                Name = accountRepository.AccountName + " Workspace"
            };
            var workspace = await docuSignApiProvider.Workspace2.CreateWorkspaceAsync(accountRepository.AccountId, workspaceBody);
            physiciansWorkspaces.Add(new PhysicianModel
            {
                Name = accountRepository.AccountName,
                WorkspaceId = workspace.WorkspaceId
            });
        }

        return physiciansWorkspaces;
    }

    public async Task<List<CareDocumentsModel>> SubmitToPhysician(SubmitToPhysiciansModel model)
    {
        const string sentStatus = "sent";
        var documents = new List<CareDocumentsModel>();

        var userForCreate = new WorkspaceUserForCreate
        {
            Email = model.Email,
            FirstName = model.Physician.Name,
            LastName = ""
        };
        await docuSignApiProvider.WorkspaceUsers.AddWorkspaceUserAsync(accountRepository.AccountId, model.Physician.WorkspaceId, userForCreate);

        foreach (var document in model.Documents)
        {
            var documentRequest = new AddWorkspaceDocumentRequest
            {
                File = new AddWorkspaceDocumentRequestFile
                {
                    Content = Convert.FromBase64String(document.Base64String),
                    FileName = document.Name
                }
            };
            var addedDocument = await docuSignApiProvider.WorkspaceDocuments.AddWorkspaceDocumentAsync(accountRepository.AccountId, model.Physician.WorkspaceId, documentRequest);

            if (document.IsForSignature)
            {
                var workspaceEnvelopeForCreate = new WorkspaceEnvelopeForCreate
                {
                    EnvelopeName = document.Name,
                    DocumentIds = [addedDocument.DocumentId]
                };

                var envelopeResponse = await docuSignApiProvider.Workspace2.CreateWorkspaceEnvelopeAsync(
                    accountRepository.AccountId,
                    model.Physician.WorkspaceId,
                    workspaceEnvelopeForCreate);

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

                var signer1 = new Signer
                {
                    Name = model.Physician.Name,
                    Email = model.Email,
                    RecipientId = "1",
                    Status = sentStatus,
                    Tabs = signer1Tabs
                };

                var recipients = new Recipients
                {
                    Signers = [signer1]
                };

                var updateEnvelopeRequest = new Envelope
                {
                    EmailSubject = "Please Sign",
                    EmailBlurb = "Sample text for email body",
                    EnvelopeId = envelopeResponse.EnvelopeId,
                    Recipients = recipients,
                    Status = sentStatus,
                };
                await docuSignApiProvider.EnvelopApi.UpdateAsync(accountRepository.AccountId, envelopeResponse.EnvelopeId, updateEnvelopeRequest);

                documents.Add(new CareDocumentsModel(document.Name, document.IsForSignature, sentStatus));
            }
            else
            {
                documents.Add(new CareDocumentsModel(document.Name, document.IsForSignature, string.Empty));
            }
        }

        return documents;
    }
}
