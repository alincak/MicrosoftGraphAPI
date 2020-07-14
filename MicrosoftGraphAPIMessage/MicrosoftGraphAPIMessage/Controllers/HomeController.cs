using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Integrations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MicrosoftGraphAPIMessage.Models;

namespace MicrosoftGraphAPIMessage.Controllers
{
  public class HomeController : Controller
  {
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
      _logger = logger;
    }

    public IActionResult Index()
    {
      return View();
    }

    public IActionResult Privacy()
    {
      return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
      return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    [HttpPost]
    public async Task<IActionResult> SendMail()
    {
      var files = Request.Form.Files;

      var _token = "";
      var microsoftGraph = new MicrosoftGraph(_token);

      var toEmails = new List<string> { "email@domain.com" };
      var messageId = await microsoftGraph.CreateDraftMessage(null, toEmails, "Medium test", "medium body");
      if (!string.IsNullOrEmpty(messageId))
      {
        var attachmentsResult = await AddAttachmentToOffice365Message(microsoftGraph, files, messageId);
        if (!attachmentsResult)
        {
          return Content("Failed.");
        }

        var sendResult = await microsoftGraph.SendMessage(messageId);
        if (!sendResult.success)
        {
          return Content("Failed.");
        }
      }

      return Content("OK.");
    }

    //Aslında burası service içerisinde ama burada öyle bir yapı olmadığından buraya ekledim.
    private async Task<bool> AddAttachmentToOffice365Message(MicrosoftGraph microsoftGraph, IFormFileCollection files, string messageId)
    {
      if (microsoftGraph == null || files == null)
      {
        return false;
      }

      foreach (var file in files)
      {
        var result = await microsoftGraph.AddAttachmentToMessage(messageId, file.FileName, IFormFileToArray(file));
        if (!result)
        {
          return false;
        }
      }

      return true;
    }

    //Aslında burasıda Helper içerisinde
    private byte[] IFormFileToArray(IFormFile file)
    {
      using (var memoryStream = new MemoryStream())
      {
        file.CopyTo(memoryStream);
        return memoryStream.ToArray();
      }
    }

  }
}
