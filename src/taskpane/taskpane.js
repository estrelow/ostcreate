/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {

  const sHost=Office.context.roamingSettings.get('host');

  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("logo").src=sHost+'/scp/logo.php?login';
    document.getElementById("run").onclick = run;
    document.getElementById("config-button").onclick = ShowConfig;
    document.getElementById("config-save").onclick = SaveConfig;
   

  }
});

/***************************************************
 * ShowConfig()
 * 
 * Show the Config <div> that's otherwise hidden
 ***************************************************/
export async function ShowConfig() {
   document.getElementById("hostname").value=Office.context.roamingSettings.get('host');
   document.getElementById("config-pane").style.display="flex";
}

/***************************************************
 * SaveConfig()
 * 
 * Saves the config values to the roamingSettings
 * structure
 ***************************************************/
export async function SaveConfig() {
   const sHost=document.getElementById("hostname").value;
   Office.context.roamingSettings.set('host',sHost);
   Office.context.roamingSettings.set('api_key',document.getElementById("api-key").value);
   document.getElementById("logo").src=sHost+'/scp/logo.php?login';
   document.getElementById("config-pane").style.display="none";
}

/*************************************************
 * run()
 * 
 * Sends the message contents to the osTicket API
 * in order to create a support ticket
 ************************************************/
export async function run() {


  const item = Office.context.mailbox.item;
  const ticket = {};
  const api_key=Office.context.roamingSettings.get('api_key');
  const sHost=Office.context.roamingSettings.get('host');

  const url=sHost+"/api/tickets.json";

  ticket.email=item.from.emailAddress;
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
  // console.log(sMsg);
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
