/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

function onNewMessageComposeHandler(event) {
  setSubject(event);
}

function setSubject(event) {
  Office.context.mailbox.item.subject.setAsync(
    "Set by an event-based add-in!",
    {
      "asyncContext": event
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() to signal to the Outlook client that the add-in has completed processing the event.
      asyncResult.asyncContext.completed();
    });
}

function onMessageRecipientsChangedHandler(event) {
  const recipientFields = event.changedRecipientFields;
  let recipientField;
  if (recipientFields.to) {
    recipientField = Office.context.mailbox.item.to;
    recipientField.getAsync({asyncContext: { event: event, recipientField: "To" }}, onRecipientsReceived);
  } else if (recipientFields.cc) {
    recipientField = Office.context.mailbox.item.cc;
    recipientField.getAsync({asyncContext: { event: event, recipientField: "Cc" }}, onRecipientsReceived);
  } else if (recipientFields.bcc) {
    recipientField = Office.context.mailbox.item.bcc;
    recipientField.getAsync({asyncContext: { event: event, recipientField: "Bcc" }}, onRecipientsReceived);
  }
}

function onRecipientsReceived(result) {
  const event = result.asyncContext.event;
  const recipientField = result.asyncContext.recipientField;
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.log(`Failed to get recipients from ${recipientField} field: ${result.error.message}`);
    event.completed();
    return;
  }

  const recipients = JSON.stringify(result.value);
  const keyName = `tagExternal${recipientField}`;
  console.log(`KEYNAME: ${keyName}`);
  if (recipients != null
    && recipients.length > 0
    && recipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
    _setSessionData(keyName, true, event);
  } else {
    _setSessionData(keyName, false, event);
  }
}

/**
 * Sets the value of the specified sessionData key.
 * If value is true, also tag as external, else check entire sessionData property bag.
 * @param {string} key The key or name
 * @param {bool} value The value to assign to the key
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
 function _setSessionData(key, value, event) {
  console.log(`SAMPLE: _setSessionData called with key=${key}, value=${value}`);
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    { asyncContext: event },
    (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(`Failed to set ${key} sessionData to ${value}. Error: ${result.error.message}`);
        event.completed();
        return;
      }

      console.log(`Set sessionData (${key}) to ${value} successfully.`);
      if (value) {
        _tagExternal(value, event);
      } else {
        _checkForExternal(event);
      }
    }
  );
}

/**
 * If there are any external recipients, prepends the subject of the Outlook item
 * with "[External]" and appends a disclaimer to the item body. If there are
 * no external recipients, ensures the tag isn't present and clears the disclaimer.
 * @param {bool} hasExternal If there are any external recipients
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function _tagExternal(hasExternal, event) {
  console.log(`SAMPLE: _tagExternal called with hasExternal=${hasExternal}.`);
  const externalTag = "[External]";

  if (hasExternal) {
    // Ensure "[External]" is prepended to the subject.
    Office.context.mailbox.item.subject.getAsync(
      { asyncContext: event },
      result => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log(`Failed to get subject: ${result.error.message}`);
          event.completed();
          return;
        }

        let subject = result.value;
        if (!subject.includes(externalTag)) {
          subject = `${externalTag} ${subject}`;
          Office.context.mailbox.item.subject.setAsync(
            subject,
            { asyncContext: event },
            result => {
              const event = result.asyncContext;
              if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(`Failed to set Subject: ${result.error.message}`);
                event.completed();
                return;
              }
              event.completed();
            }
          );
        }
      }
    );
  } else {
    // Ensure "[External]" is not part of the subject.
    Office.context.mailbox.item.subject.getAsync(
      { asyncContext: event },
      result => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log(`Failed to get subject: ${result.error.message}`);
          event.completed();
          return;
        }

        const currentSubject = result.value;
        if (currentSubject.startsWith(externalTag)) {
          const updatedSubject = currentSubject.replace(externalTag, "");
          const subject = updatedSubject.trim();
          Office.context.mailbox.item.subject.setAsync(
            subject,
            { asyncContext: event },
            result => {
              const event = result.asyncContext;
              if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(`Failed to set subject: ${result.error.message}`);
                event.completed();
                return;
              }
              event.completed();
            }
          );
        }
      }
    );
  }
}

/**
 * Checks the sessionData property bag to determine if any field contains external recipients.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function _checkForExternal(event) {
  // Get sessionData to determine if any fields have external recipients.
  Office.context.mailbox.item.sessionData.getAllAsync(
    { asyncContext: event },
    result => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(`Failed to get all sessionData: ${result.error.message}`);
        event.completed();
        return;
      }

      const sessionData = JSON.stringify(result.value);
      if (sessionData != null
        && sessionData.length > 0
        && sessionData.includes("true")) {
        _tagExternal(true, event);
      } else {
        _tagExternal(false, event);
      }
    }
  );
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);