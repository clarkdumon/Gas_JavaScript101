  let EMAILto = ['toEmail']
  let EMAILsubject ='Automated PTO Email | '
  const HTMLOptions = {
    from: '--fromEmail',
    name: 'PTO AUTOMATED EMAIL',
    replyTo:'--Reply to Email',
  }
const SIGNATURE = `<div>
    <br />
    <b style="color: rgb(0, 0, 0); font-family: 'Century Gothic', serif;font-size: 14px;text-transform: uppercase;">WORKFORCE MANAGEMENT</b>
    <span>&#8211; </span>
    <b style="color: rgb(117, 117, 117);font-family: 'Century Gothic', serif;font-size: 12px;text-transform: uppercase;">SCHEDULING TEAM</b>
    <br>
    <span style="margin:0;font-family: 'Century Gothic', serif;font-size: 12px;line-height: 1.2;font-weight: bold;color:rgb(102, 102, 102)">PTO AUTOMATED EMAIL</span>
    <br />
    <br />
    <img
      src="https://ci3.googleusercontent.com/meips/ADKq_NbQwD5g9y4c7nCFJG5Dk7LMOP3yhWGcDzgn3qPR8OCWl4maIa-ZeEy7u46h6tcaFs6LdIc2dsizdHCi1R229Jkft734ykqPUWpv6DccNmID_dFBSyDqVTIU=s0-d-e1-ft#https://seeklogo.com/images/T/tdcx-logo-4AC80AB1A3-seeklogo.com.png"
      width="96"
      height="39" />
    <br />
    <span style="font-family: 'Century Gothic', serif;color: rgb(233, 29, 36);line-height: 1;font-size: 12px;">w.</span>
    <span style="background-color: transparent;font-family: 'Century Gothic', serif;line-height: 1;font-size: 12px;">www.tdcx.com</span>
    <br />
    <span style="font-family: 'Century Gothic', serif;font-size: 11px;line-height: 1.2;">
      This is a confidential email that may be privileged or legally protected.
      You are not authorized to copy or disclose the contents of this email. If
      you are not the intended addressee, please inform the sender and delete
      this email.
    </span>
  </div>`



function emailPTO() {

  EMAILsubject = EMAILsubject + 'Approved'
  HTMLOptions.htmlBody =
    `
  Hi [NAME],
  <br>
  <br>
  Your PTO has been approved for [LEAVE DATE]
  <br>
  Kindly file your PTO in flash to avoid pay discrepancies
  <br><br>
${SIGNATURE}`


  GmailApp.sendEmail(EMAILto, EMAILsubject, null, HTMLOptions)
}

function deniedemailPTO() {

HTMLOptions.htmlBody=
    `
  Hi [NAME],
  <br><br>
    Your PTO has been Denied due to [Remarks for denial] for [LEAVE DATE]
  <br>
    Kindly request you PTO in flash
  <br><br><br><br>
${SIGNATURE}`

GmailApp.sendEmail(EMAILto, EMAILsubject, null, HTMLOptions)

}
