const { EmailClient, KnownEmailSendStatus } = require("@azure/communication-email");
require('dotenv').config()
// https://learn.microsoft.com/en-us/azure/communication-services/
// https://learn.microsoft.com/en-us/azure/communication-services/quickstarts/email/send-email-smtp/send-email-smtp?pivots=smtp-method-smtpclient
const connectionString = process.env.CONNECTION_STRING;
const senderAddress = process.env.SENDER_ADDRESS
const recipientAddress = process.env.RECIPIENT_ADDRESS

async function main() {
  const POLLER_WAIT_TIME = 10

  const message = {
    senderAddress: senderAddress,
    recipients: {
      to: [{ address: recipientAddress }],
    },
    content: {
      subject: "Test email from JS Sample",
      plainText: "This is plaintext body of test email.",
      html: "<html><h1>Hello World...</h1></html>",
    },
  }

  try {
    const client = new EmailClient(connectionString);

    const poller = await client.beginSend(message);

    if (!poller.getOperationState().isStarted) {
      throw "Poller was not started."
    }

    let timeElapsed = 0;
    while(!poller.isDone()) {
      poller.poll();
      console.log("Email send polling in progress");

      await new Promise(resolve => setTimeout(resolve, POLLER_WAIT_TIME * 1000));
      timeElapsed += 10;

      if(timeElapsed > 18 * POLLER_WAIT_TIME) {
        throw "Polling timed out.";
      }
    }

    if(poller.getResult().status === KnownEmailSendStatus.Succeeded) {
      console.log(`Successfully sent the email (operation id: ${poller.getResult().id})`);
    }
    else {
      throw poller.getResult().error;
    }
  }
  catch(ex) {
    console.error(ex);
  }
}

main();
