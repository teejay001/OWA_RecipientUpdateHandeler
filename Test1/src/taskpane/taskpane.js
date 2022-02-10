/* eslint-disable no-debugger */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = printFirstRecipientEmail;
    addMailboxItemEventAsync(Office.EventType.RecipientsChanged, handlerRecipientsChanged);
  }
});

const handlerRecipientsChanged = (args) => {
  console.log(args);
  printFirstRecipientEmail();
};

async function printFirstRecipientEmail() {
  const toRecipients = await getRecipientsAsync();

  if (toRecipients.length > 0) {
    document.getElementById("item-recipient").innerHTML =
      "<b>First Recipient Email:</b> <br/>" + toRecipients[0].emailAddress;
  } else {
    document.getElementById("item-recipient").innerHTML = "<b>No email address</b>";
  }
}

async function addMailboxItemEventAsync(event, handler) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.addHandlerAsync(event, handler, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        reject(asyncResult.error);
      } else {
        resolve();
      }
    });
  });
}

async function getRecipientsAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.to.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error);
        reject();
      } else {
        resolve(asyncResult.value);
      }
    });
  });
}
