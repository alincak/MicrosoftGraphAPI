using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Integrations
{
  public class MicrosoftGraph
  {
    private readonly GraphServiceClient graphClient;

    public MicrosoftGraph(string token)
    {
      graphClient = CreateGraphClientAsync(token);
    }

    private GraphServiceClient CreateGraphClientAsync(string token)
    {
      return new GraphServiceClient(
             new DelegateAuthenticationProvider(
                (request) =>
                {
                  if (!string.IsNullOrEmpty(token))
                  {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                  }
                  return Task.FromResult(0);
                }));
    }

    public async Task<string> CreateDraftMessage(string messageId, IList<string> toEmails, string subject, string body)
    {
      if (toEmails == null || !toEmails.Any() || string.IsNullOrEmpty(subject) || string.IsNullOrEmpty(body))
      {
        return null;
      }

      var message = new Message
      {
        Subject = subject,
        Body = new ItemBody
        {
          ContentType = BodyType.Html,
          Content = body
        },
        ToRecipients = toEmails.Select(x => new Recipient { EmailAddress = new EmailAddress { Address = x } })
      };

      try
      {
        Message resultMessage = null;
        if (string.IsNullOrEmpty(messageId))
        {
          resultMessage = await graphClient.Me.Messages
                                .Request()
                                .AddAsync(message);
        }
        else
        {
          resultMessage = await graphClient.Me.Messages[messageId]
                                 .CreateReply(message)
                                 .Request()
                                 .PostAsync();
        }

        return resultMessage?.Id;
      }
      catch (Exception ex)
      {
        //Logger.ErrorFormat("CreateDraftMessage - {0}", ex);
        return null;
      }
    }

    public async Task<bool> AddAttachmentToMessage(string messageId, string fileName, byte[] bytes)
    {
      try
      {
        var result = true;
        if (bytes.Length >= 3000000)
        {
          result = await MessageCreateUploadSession(messageId, fileName, bytes);
        }
        else
        {
          var attachment = new FileAttachment
          {
            Name = fileName,
            ContentBytes = bytes
          };

          await graphClient.Me.Messages[messageId].Attachments
                      .Request()
                      .AddAsync(attachment).ConfigureAwait(false);
        }

        if (!result) throw new Exception("Upload failed.");

        return result;
      }
      catch
      {
        await DeleteMessage(messageId);
        return false;
      }
    }

    private async Task<bool> MessageCreateUploadSession(string messageId, string fileName, byte[] bytes)
    {
      try
      {
        using (Stream stream = new MemoryStream(bytes))
        {
          var attachmentItem = new AttachmentItem
          {
            AttachmentType = AttachmentType.File,
            Name = fileName,
            Size = stream.Length
          };

          var uploadSession = await graphClient.Me.Messages[messageId].Attachments
                                    .CreateUploadSession(attachmentItem)
                                    .Request()
                                    .PostAsync();

          var maxChunkSize = 320 * 1024;
          var largeFileUploadTask = new LargeFileUploadTask<FileAttachment>(uploadSession, stream, maxChunkSize);

          var uploadResult = await largeFileUploadTask.UploadAsync();
          if (uploadResult.UploadSucceeded)
          {
            return true;
          }

          return false;
        }
      }
      catch (Exception ex)
      {
        //Logger.ErrorFormat("CreateUploadSession - Upload failed ex: {0}", ex);
        return false;
      }
    }

    public async Task<(bool success, string message)> SendMessage(string messageId)
    {
      try
      {
        await graphClient.Me.Messages[messageId]
                    .Send()
                    .Request()
                    .PostAsync();

        return (true, null);
      }
      catch (Exception ex)
      {
        //Logger.ErrorFormat("SendMessage - {0}", ex);
        await DeleteMessage(messageId);
        return (false, ex.Message);
      }
    }

    private async Task DeleteMessage(string messageId)
    {
      try
      {
        await graphClient.Me.Messages[messageId]
              .Request()
              .DeleteAsync();
      }
      catch (Exception ex)
      {
        //Logger.ErrorFormat("DeleteMessage - ex: {0}", ex);
      }
    }

  }
}
