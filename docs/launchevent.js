/*
 * launchevent.js
 *
 * Smart Alerts add-in: makes sure tim@blob.be is a recipient (To, Cc or Bcc)
 * before a message is sent.
 *
 * Flow:
 *   1. OnMessageSend fires -> onMessageSendHandler reads To, Cc and Bcc.
 *   2. Tim is present                 -> event.completed({ allowEvent: true }), no dialog.
 *   3. Tim is missing                 -> event.completed({ allowEvent: false, ... }) with
 *        - a "Take Action" button ("Tim toevoegen in cc") that runs addTimToCcCommand
 *          (Mailbox 1.15: function command from the Smart Alerts dialog),
 *        - a "Send Anyway" button, enabled at runtime with sendModeOverride = promptUser
 *          (Mailbox 1.14), because the manifest send mode is SoftBlock.
 *   4. Recipients cannot be read      -> fail open: the message is sent (see ALLOW_SEND_ON_ERROR).
 *
 * Notes:
 *   - No imports and no modules: classic Outlook on Windows loads this file directly
 *     in a JavaScript-only runtime, so everything lives in this single file.
 *   - Plain ES5 on purpose, for the same reason.
 *   - Office.onReady / Office.initialize do NOT run when an event handler is launched,
 *     so all start-up logic lives inside the handlers.
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

var TIM_ADDRESS = "tim@blob.be";
var TIM_DISPLAY_NAME = "Tim";

// User-facing texts (kept as constants so they are easy to change).
var MSG_TIM_MISSING = "Tim staat niet in de ontvangers.";
var MSG_TIM_MISSING_NO_ACTION =
  "Tim staat niet in de ontvangers. Voeg hem toe via de knop \"Tim toevoegen in cc\" in het lint, of verstuur toch zonder Tim.";
var MSG_ADD_FAILED = "Tim kon niet in Cc worden gezet. Voeg hem handmatig toe.";
var MSG_SEND_FAILED = "Tim is in Cc gezet, maar het verzenden is mislukt. Klik zelf op Verzenden.";

// Smart Alerts limits the custom button text to 20 characters.
var ADD_TIM_BUTTON_LABEL = "Tim toevoegen in cc";

// Must match the id of the <Control> element in manifest.xml.
var ADD_TIM_COMMAND_ID = "addTimToCcButton";

// Marker passed through contextData so the function command knows it was started
// from the Smart Alerts dialog (and not from the ribbon button).
var DIALOG_CONTEXT_SOURCE = "smartAlertsDialog";

// If reading the recipients fails or hangs, let the message through instead of blocking it.
var ALLOW_SEND_ON_ERROR = true;

// Smart Alerts shows a "taking longer than expected" dialog after 5 seconds.
// Stay below that so the user never sees it because of this add-in.
var RECIPIENTS_TIMEOUT_MS = 4000;

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function log(message, detail) {
  try {
    if (typeof console !== "undefined" && console.log) {
      if (detail !== undefined) {
        console.log("[TimCcCheck] " + message, detail);
      } else {
        console.log("[TimCcCheck] " + message);
      }
    }
  } catch (e) {
    // Logging must never break the handler.
  }
}

function isMailboxSetSupported(version) {
  try {
    return Office.context.requirements.isSetSupported("Mailbox", version);
  } catch (e) {
    return false;
  }
}

function getPromptUserOverride() {
  try {
    if (Office.MailboxEnums && Office.MailboxEnums.SendModeOverride) {
      return Office.MailboxEnums.SendModeOverride.PromptUser;
    }
  } catch (e) {
    // Fall through to the literal value.
  }
  return "promptUser";
}

function normalizeAddress(address) {
  return String(address || "").replace(/^\s+|\s+$/g, "").toLowerCase();
}

function isTim(address) {
  return normalizeAddress(address) === normalizeAddress(TIM_ADDRESS);
}

/**
 * Reads one recipient field (to, cc or bcc) of the item being composed.
 * Calls back with (error, arrayOfEmailAddresses). The callback is called exactly once.
 */
function getRecipientAddresses(recipientField, callback) {
  var called = false;
  function done(error, addresses) {
    if (called) {
      return;
    }
    called = true;
    callback(error, addresses);
  }

  try {
    if (!recipientField || typeof recipientField.getAsync !== "function") {
      done(new Error("Recipient field is not available."), []);
      return;
    }
    recipientField.getAsync(function (result) {
      try {
        if (!result || result.status !== Office.AsyncResultStatus.Succeeded) {
          done(new Error((result && result.error && result.error.message) || "getAsync failed."), []);
          return;
        }
        var addresses = [];
        var recipients = result.value || [];
        for (var i = 0; i < recipients.length; i++) {
          if (recipients[i] && recipients[i].emailAddress) {
            addresses.push(recipients[i].emailAddress);
          }
        }
        done(null, addresses);
      } catch (e) {
        done(e, []);
      }
    });
  } catch (e) {
    done(e, []);
  }
}

/**
 * Reads To, Cc and Bcc in parallel and calls back with (error, allAddresses).
 * If any of the three fails, the error is reported and the result is not trusted.
 */
function getAllRecipientAddresses(callback) {
  var item;
  try {
    item = Office.context.mailbox.item;
  } catch (e) {
    callback(e, []);
    return;
  }
  if (!item) {
    callback(new Error("No item is available."), []);
    return;
  }

  var fields = [item.to, item.cc, item.bcc];
  var pending = fields.length;
  var firstError = null;
  var all = [];

  for (var i = 0; i < fields.length; i++) {
    getRecipientAddresses(fields[i], function (error, addresses) {
      if (error && !firstError) {
        firstError = error;
      }
      all = all.concat(addresses);
      pending--;
      if (pending === 0) {
        callback(firstError, all);
      }
    });
  }
}

function containsTim(addresses) {
  for (var i = 0; i < addresses.length; i++) {
    if (isTim(addresses[i])) {
      return true;
    }
  }
  return false;
}

function showErrorNotification(key, message) {
  try {
    Office.context.mailbox.item.notificationMessages.addAsync(
      key,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: message
      },
      function () {
        // Best effort only.
      }
    );
  } catch (e) {
    log("Could not show notification.", e);
  }
}

// ---------------------------------------------------------------------------
// OnMessageSend handler (Smart Alerts)
// ---------------------------------------------------------------------------

function buildBlockOptions() {
  var options = {
    allowEvent: false,
    // Adds the "Send Anyway" button. The manifest send mode stays SoftBlock.
    sendModeOverride: getPromptUserOverride()
  };

  if (isMailboxSetSupported("1.15")) {
    // The Smart Alerts dialog can run a function command (Mailbox 1.15).
    options.errorMessage = MSG_TIM_MISSING;
    options.cancelLabel = ADD_TIM_BUTTON_LABEL; // Label of the "Take Action" button.
    options.commandId = ADD_TIM_COMMAND_ID;
    options.contextData = JSON.stringify({ source: DIALOG_CONTEXT_SOURCE });
  } else {
    // Mailbox 1.14 only: no function command from the dialog.
    // The ribbon button of the add-in still adds Tim to Cc.
    options.errorMessage = MSG_TIM_MISSING_NO_ACTION;
  }

  return options;
}

// DIAGNOSTIC SWITCH. When true, every message is let through immediately, without reading
// recipients. Use it to find out whether the handler runs at all. Set back to false afterwards.
var DEBUG_PASS_THROUGH = true;

function onMessageSendHandler(event) {
  if (DEBUG_PASS_THROUGH) {
    event.completed({ allowEvent: true });
    return;
  }
  var finished = false;
  var timer = null;

  // Completes the event exactly once, whatever happens.
  function finish(options) {
    if (finished) {
      return;
    }
    finished = true;
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
    try {
      event.completed(options);
    } catch (e) {
      log("event.completed failed.", e);
    }
  }

  function failSafeOptions() {
    if (ALLOW_SEND_ON_ERROR) {
      return { allowEvent: true };
    }
    return {
      allowEvent: false,
      errorMessage: "De ontvangers konden niet worden gecontroleerd.",
      sendModeOverride: getPromptUserOverride()
    };
  }

  try {
    // Guard against a getAsync call that never returns.
    timer = setTimeout(function () {
      log("Reading the recipients timed out.");
      finish(failSafeOptions());
    }, RECIPIENTS_TIMEOUT_MS);

    getAllRecipientAddresses(function (error, addresses) {
      try {
        if (containsTim(addresses)) {
          // Tim is a recipient: send without any dialog. This also covers the case where
          // only one of the three lookups failed but Tim was found anyway.
          finish({ allowEvent: true });
          return;
        }
        if (error) {
          log("Could not read the recipients.", error);
          finish(failSafeOptions());
          return;
        }
        finish(buildBlockOptions());
      } catch (e) {
        log("Unexpected error while evaluating the recipients.", e);
        finish(failSafeOptions());
      }
    });
  } catch (e) {
    log("Unexpected error in onMessageSendHandler.", e);
    finish(failSafeOptions());
  }
}

// ---------------------------------------------------------------------------
// Function command: "Tim toevoegen in cc"
// Runs from the Smart Alerts "Take Action" button (Mailbox 1.15) and from the ribbon button.
// ---------------------------------------------------------------------------

/**
 * Calls back with true if the command was started from the Smart Alerts dialog.
 * Any problem is treated as "started from the ribbon".
 */
function wasStartedFromDialog(callback) {
  var called = false;
  function done(value) {
    if (called) {
      return;
    }
    called = true;
    callback(value);
  }

  try {
    var item = Office.context.mailbox.item;
    if (!item || typeof item.getInitializationContextAsync !== "function") {
      done(false);
      return;
    }
    item.getInitializationContextAsync(function (result) {
      try {
        if (!result || result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
          done(false);
          return;
        }
        // Older clients return a string, newer ones can return an object.
        var data = typeof result.value === "string" ? JSON.parse(result.value) : result.value;
        done(!!data && data.source === DIALOG_CONTEXT_SOURCE);
      } catch (e) {
        done(false);
      }
    });
  } catch (e) {
    done(false);
  }
}

/** Adds Tim to Cc unless he is already a recipient. Calls back with (error). */
function ensureTimInCc(callback) {
  getAllRecipientAddresses(function (readError, addresses) {
    if (!readError && containsTim(addresses)) {
      callback(null);
      return;
    }
    // If reading failed we still try to add him; a duplicate recipient is harmless.
    try {
      Office.context.mailbox.item.cc.addAsync(
        [{ displayName: TIM_DISPLAY_NAME, emailAddress: TIM_ADDRESS }],
        function (result) {
          if (result && result.status === Office.AsyncResultStatus.Succeeded) {
            callback(null);
          } else {
            callback(new Error((result && result.error && result.error.message) || "cc.addAsync failed."));
          }
        }
      );
    } catch (e) {
      callback(e);
    }
  });
}

function addTimToCcCommand(event) {
  var finished = false;
  function finish() {
    if (finished) {
      return;
    }
    finished = true;
    try {
      event.completed();
    } catch (e) {
      log("event.completed failed in addTimToCcCommand.", e);
    }
  }

  try {
    wasStartedFromDialog(function (fromDialog) {
      ensureTimInCc(function (error) {
        if (error) {
          log("Could not add Tim to Cc.", error);
          showErrorNotification("timCcAddFailed", MSG_ADD_FAILED);
          finish();
          return;
        }

        var item = Office.context.mailbox.item;
        if (!fromDialog || !item || typeof item.sendAsync !== "function") {
          // Started from the ribbon (or sendAsync is unavailable): the user sends the message.
          finish();
          return;
        }

        // Started from the Smart Alerts dialog: continue sending now that Tim is in Cc.
        // OnMessageSend fires again and passes silently because Tim is a recipient.
        try {
          item.sendAsync(function (result) {
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) {
              log("sendAsync failed.", result && result.error);
              showErrorNotification("timCcSendFailed", MSG_SEND_FAILED);
            }
            finish();
          });
        } catch (e) {
          log("sendAsync threw.", e);
          showErrorNotification("timCcSendFailed", MSG_SEND_FAILED);
          finish();
        }
      });
    });
  } catch (e) {
    log("Unexpected error in addTimToCcCommand.", e);
    finish();
  }
}

// ---------------------------------------------------------------------------
// IMPORTANT: map the function names used in manifest.xml to their JavaScript functions.
// Without this, Outlook blocks sending (SoftBlock) because the handler cannot be found.
// ---------------------------------------------------------------------------

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("addTimToCcCommand", addTimToCcCommand);
