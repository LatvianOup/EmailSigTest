Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
      // Function to insert a signature when the add-in is triggered
      window.insertSignature = async function() {
        const item = Office.context.mailbox.item;
        const body = item.body;
  
        // Insert the 'Hello World' signature at the end of the message
        const signature = "<p>Hello World</p>";
        
        // Insert the signature at the bottom of the current email body
        await body.setAsync(body + signature, { coercionType: Office.CoercionType.Html });
      };
    }
  });
  