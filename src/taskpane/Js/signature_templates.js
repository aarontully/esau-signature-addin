const { is } = require("core-js/core/object");

function get_template_A_str(user_info)
{
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
  return str;
}

function get_template_B_str(user_info)
{
  let str = `
    <div>
      ${is_valid_data(user_info.greeting) ? `<div style="padding-bottom: 16px;">${user_info.greeting}</div>` : ""}
      <div style="color: #FF4370;"><strong>${user_info.name} ${is_valid_data(user_info.pronoun) ? `(${user_info.pronoun})` : ""}</strong></div>
      <div style="padding-bottom: 16px;">${user_info.title} ${is_valid_data(user_info.department) ? ` | ${user_info.department}` : ""}</div>
    </div>
  `;
  return str;
}