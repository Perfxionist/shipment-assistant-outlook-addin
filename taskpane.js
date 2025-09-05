
Office.onReady(() => {
  if (Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;

    item.body.getAsync("text", { asyncContext: "body" }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const bodyText = result.value;

        const shipmentRefs = bodyText.match(/(?:HBL|MBL|PO)[\s\-:]?\w+/gi) || [];

        document.getElementById("shipmentRefs").innerText = shipmentRefs.join("\n");

        if (shipmentRefs.length > 0) {
          document.getElementById("actions").innerHTML = `
            <button onclick="openCW1('${shipmentRefs[0]}')">Open in CW1</button>
          `;
        } else {
          document.getElementById("actions").innerHTML = `
            <button onclick="triggerOCR()">Create Shipment via OCR</button>
          `;
        }
      }
    });
  }
});

function openCW1(ref) {
  const cw1Url = `https://cw1.yourdomain.com/open?ref=${encodeURIComponent(ref)}`;
  window.open(cw1Url, "_blank");
}

function triggerOCR() {
  alert("OCR triggered. Document parsing in progress...");
}
