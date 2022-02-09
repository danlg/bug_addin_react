console.log("commands.js called")
Office.onReady(() => {
    console.log("Office.onReady success") // If needed, Office.js is ready to be called
});

//======
// require('dotenv').config();
//export async function writeToWordImpl (text, evt) {
let count = 0
export async function writeToWordImpl (text) {
    return Word.run(async context => {
        count = count +1
        let text_ = text  + count.toString()
        console.log("Writing Word:", text_)
        // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end)
        // Create a proxy object for the document body.
        var body = context.document.body;
        console.log("Writing Word - BODY:", text_)
        // Queue a command to insert the paragraph at the end of the document body.
        // const paragraph = body.insertParagraph(text, Word.InsertLocation.end)
        //paragraph.font.color = "black"
        try{
            console.log("Writing Word - B4 TRY:", text_)
            body.insertParagraph(text_, Word.InsertLocation.end)
            console.log("Writing Word - SHOULD HAVE INSERTED OK:", text_ )
        }
        catch(error) {
            //never printed
            console.log("Writing Word - CATCH:", error)
        }
        console.log("Writing Word - SHOULD HAVE INSERTED OK 2:", text_)
        // body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end)
        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync().then(function () {
            console.log('Paragraph added at the end of the document body.');
        })
        .catch(function (error) {
            //never printed
            console.log('Error sync paragragh: ', error);
        })
        console.log("Writing to Word 8 - SYNC:", text)
    })
    .catch(function (error) {
        //never printed
        console.log('Error insert paragragh: ', error);
    })
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
async function action(evt) {
//   // Show a notification message
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
    console.log("INSIDE action X ")
    await writeToWordImpl("BeFoRE displayDialogAsync")
    Office.context.ui.displayDialogAsync('https://localhost:3000/dialog.html', {
            height: 20,
            width: 20,
            // displayInIframe: true
        },
        async function (asyncResult) {
            dialog = asyncResult.value;
            //await writeToWord("iNsIDE displayDialogAsync")
            console.log("displayDialogAsync")
            // dialog.addEventHandler(
            //     processMessage,
            //     Office.EventType.DialogMessageReceived, processMessage);
        }
    );
    // evt.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
