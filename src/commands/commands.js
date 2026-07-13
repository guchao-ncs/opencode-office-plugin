/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, localStorage, console */

Office.onReady(() => {
  // Ribbon command handler stub. The actual AI Assistant functionality is in taskpane.html/js.
});

/**
 * Ribbon command handler for the Task Pane button.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // The actual "Show Task Pane" action is wired via manifest.xml's
  // ShowTaskpane action, not this handler function.
  event.completed();
}

// Register the function with Office (required even though ShowTaskpane
// is handled declaratively in manifest.xml).
Office.actions.associate("action", action);

/**
 * Context menu handler to analyze the selected text.
 * @param event {Office.AddinCommands.Event}
 */
function analyzeSelectionAction(event) {
  if (typeof Office !== "undefined" && Office.context) {
    if (Office.context.document) {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const text = asyncResult.value;
            if (text && text.trim()) {
              localStorage.setItem(
                "contextMenuTrigger",
                JSON.stringify({
                  action: "analyze",
                  content: text.trim(),
                  timestamp: Date.now(),
                })
              );
              Office.addin.showAsTaskpane().catch((err) => {
                console.error("Failed to show task pane:", err);
              });
            }
          }
          event.completed();
        }
      );
      return;
    } else if (Office.context.mailbox && Office.context.mailbox.item) {
      Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const text = asyncResult.value;
          if (text && text.trim()) {
            localStorage.setItem(
              "contextMenuTrigger",
              JSON.stringify({
                action: "analyze",
                content: text.trim(),
                timestamp: Date.now(),
              })
            );
            Office.addin.showAsTaskpane().catch((err) => {
              console.error("Failed to show task pane:", err);
            });
          }
        }
        event.completed();
      });
      return;
    }
  }
  event.completed();
}

Office.actions.associate("analyzeSelectionAction", analyzeSelectionAction);
