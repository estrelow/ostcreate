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

  const item = Office.context.mailbox.item;
  const ticket = {};
  const api_key="C987FCC9C5A4CCAD620C190514916220";
  const url="https://soporte.moller.cl/api/tickets.json";

  ticket.email="esf@moller.cl";
  ticket.subject=item.subject;
  ticket.message="Necesito una ayuda";
  ticket.name=item.from.emailAddress;

  item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Failed) {
    console.log(`Failed to get body: ${bodyResult.error.message}`);
    return;
  } else if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {

    ticket.message=bodyResult.value;
  }
  });

  const sMsg=JSON.stringify(ticket);
  console.log(sMsg);
  const hKey=new Headers();
  hKey.append('X-API-Key',api_key);
  hKey.append('Content-type','application/json');
  const options={
    method: 'POST',
    headers: hKey,
    body: JSON.stringify(ticket)
  };

  const ost=new Request(url, options);

  fetch(ost)
  .then(response => response.text())
  .then(data => {
    console.log('Success:', data);
  })
  .catch((error) => {
    console.error('Error:', error);
  });

 
}
