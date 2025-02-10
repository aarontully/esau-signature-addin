const { is } = require("core-js/core/object");

function get_template_A_str(user_info)
{
  let str = `
    <div>
      ${is_valid_data(user_info.greeting) ? `<p>${user_info.greeting}</p>` : ""}
      <p></strong>${user_info.name} ${is_valid_data(user_info.pronoun) ? `(${user_info.pronoun})` : ""}</p>
      <p>${user_info.title}</p><br/>
      ${is_valid_data(user_info.phone) ? `<p>PH: ${user_info.phone}</p>` : ""}
      <a href="https://www.endemolshine.com.au/">www.endemolshine.com.au</a>
      <img src="./assets/sig_image.png alt="logo />
      <div>This electronic mail, including any attachments, is intended for the addressee only and may contain information that is either confidential or subject to legal professional privilege. Unauthorised reproduction, use or disclosure of the contents of this mail is prohibited. If you have received this mail in error, please delete it from your system immediately and notify Endemol Shine Australia by contacting us at www.endemolshine.com.au</div>
    </div>
  `;
  return str;
}

function get_template_B_str(user_info)
{
  let str = `
    <div>
      ${is_valid_data(user_info.greeting) ? `<p>${user_info.greeting}</p>` : ""}
      <p></strong>${user_info.name} ${is_valid_data(user_info.pronoun) ? `(${user_info.pronoun})` : ""}</p>
      <p>${user_info.title}</p><br/>
    </div>
  `;
  return str;
}