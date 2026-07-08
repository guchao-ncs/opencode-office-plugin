/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

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
