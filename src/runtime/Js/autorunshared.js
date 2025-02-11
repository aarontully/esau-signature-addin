// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sig_image.png";
  let str = `
    <div>
      ${is_valid_data(user_info.greeting) ? `<div style="padding-bottom: 16px;">${user_info.greeting}</div>` : ""}
      <div style="color: #FF4370;"><strong>${user_info.name} ${is_valid_data(user_info.pronoun) ? `(${user_info.pronoun})` : ""}</strong></div>
      <div style="padding-bottom: 16px;">${user_info.title} ${is_valid_data(user_info.department) ? ` | ${user_info.department}` : ""}</div>
      ${is_valid_data(user_info.phone) ? `<div>${user_info.phone}</div>` : ""}
      <div>${address}</div>
      <div><a href="https://www.endemolshine.com.au/">www.endemolshine.com.au</a></div>
      <div><img src="./assets/sig_image.png" alt="logo" /></div>
      <div style="font-size: 10px;">This electronic mail, including any attachments, is intended for the addressee only and may contain information that is either confidential or subject to legal professional privilege. Unauthorised reproduction, use or disclosure of the contents of this mail is prohibited. If you have received this mail in error, please delete it from your system immediately and notify Endemol Shine Australia by contacting us at www.endemolshine.com.au</div>
    </div>
  `;

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "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",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = `
    <div>
      ${is_valid_data(user_info.greeting) ? `<div style="padding-bottom: 16px;">${user_info.greeting}</div>` : ""}
      <div style="color: #FF4370;"><strong>${user_info.name} ${is_valid_data(user_info.pronoun) ? `(${user_info.pronoun})` : ""}</strong></div>
      <div style="padding-bottom: 16px;">${user_info.title} ${is_valid_data(user_info.department) ? ` | ${user_info.department}` : ""}</div>
    </div>
  `;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
