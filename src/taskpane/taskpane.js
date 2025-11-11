/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

console.log("running and deployed!");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("downloadButton").onclick = downloadSampleFile;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "blue";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
  Office.onReady(() => {});
}

function downloadSampleFile() {
  console.log("[Log] Trying to download file");

  const content = "hello world";
  const blob = new Blob([content], { type: "text/plain;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  console.log("Generated Blob:", blob);
  console.log("Generated URL:", url);

  // Create a temporary link element
  const link = document.createElement("a");
  try {
    link.setAttribute("href", url);
    link.setAttribute("download", "samplefile.txt");
    link.style.visibility = "hidden";
    document.body.appendChild(link);
  } catch (error) {
    console.error("[Error] Creation of element failed", error);
  }

  // Click on link element
  try {
    link.click();
    console.log("[Log] Click event triggered");
  } catch (error) {
    console.error("[Error] Click event failed", error);
  }

  document.body.removeChild(link);

  // Clean up URL object
  URL.revokeObjectURL(url);
}
Office.onReady(() => {});
