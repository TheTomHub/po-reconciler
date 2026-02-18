/* global Office */

Office.onReady(() => {
  // Register ribbon button handlers
});

/**
 * Ribbon button handler — opens the taskpane.
 * Called when user clicks "Quick Reconcile" in the ribbon.
 */
function reconcileFromRibbon(event) {
  // Show a notification that the taskpane should be used
  Office.context.ui.displayDialogAsync(
    "https://thetomhub.github.io/po-reconciler/taskpane.html",
    { height: 60, width: 30 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Fallback: just complete the event
        event.completed();
        return;
      }
      event.completed();
    }
  );
}

// Register function with Office
Office.actions = Office.actions || {};
Office.actions.associate("reconcileFromRibbon", reconcileFromRibbon);
