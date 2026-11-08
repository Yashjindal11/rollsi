/* global document, Office, URLSearchParams, window */

Office.onReady(() => {
  const msg = new URLSearchParams(window.location.search).get("msg") ?? "";
  document.getElementById("msg")!.textContent = msg;
  document
    .getElementById("ok")!
    .addEventListener("click", () => Office.context.ui.messageParent("ok"));
});
