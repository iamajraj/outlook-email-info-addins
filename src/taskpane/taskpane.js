/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("ownerName").textContent = Office.context.mailbox.item.to[0].displayName;
    document.getElementById("ownerEmail").textContent = Office.context.mailbox.item.to[0].emailAddress;
    document.getElementById("senderName").textContent = Office.context.mailbox.item.from.displayName;
    document.getElementById("senderEmail").textContent = Office.context.mailbox.item.from.emailAddress;
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, {}, (result) => {
      if (result.error) return;
      document.getElementById("content").textContent = result.value;
    });
  }
});
