/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  var item = Office.context.mailbox.item;
  var emailAddress = Office.context.mailbox.userProfile.emailAddress;
  var itemId = item.itemId
  //  var settings = {
  //   "async": true,
  //   "crossDomain": true,
  //   "url": "https://prod-138.westus.logic.azure.com/workflows/19dea0d01b67425bb1c73965f513059f/triggers/manual/paths/invoke/messageId/"+item.itemId+"/smbName/"+emailAddress+"?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=sLBzcVeh430T8N7QQHUZIaCQ7BuJSVL9GmylzgL-Wuk",
  //   "method": "GET"
  // }
  var find = '/';
  var re = new RegExp(find, 'g');

  itemId = itemId.replace(re, '-');
  var newA = document.createElement('a');
  newA.setAttribute('href', "https://prod-138.westus.logic.azure.com/workflows/19dea0d01b67425bb1c73965f513059f/triggers/manual/paths/invoke/messageId/"+itemId+"/smbName/" + emailAddress + "?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=sLBzcVeh430T8N7QQHUZIaCQ7BuJSVL9GmylzgL-Wuk");
  newA.innerHTML = "link text";
  newA.click();

  // Write message property value to the task pane
  //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" +  item.subject+"<br/>internetMessageId:"+item.internetMessageId+"<br/>itemID:"+item.itemId+"<br/>itemID:"+item.conversationId;
  document.getElementById("item-subject").innerHTML = "<br/>itemId:" + itemId;
}
