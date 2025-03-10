/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word  */

//http-server -S -C localhost.crt -K localhost.key --cors . -p 3000

Office.onReady((info) => {  //  This mechanism checks the platform on which the add-in will be used (Word, Excel, Outlook, etc.) and initiates add-in-specific actions.
                            //  It is a readiness check provided by the Office Add-in API.
                            //  This function is called when the add-in is installed and available.
                            //  The info parameter contains information about which Office application the add-in is running on.

  if (info.host === Office.HostType.Word) { //  info.host specifies the Office application in which the add-in is running (e.g. Word, Excel, Outlook).
                                            //  Office.HostType.Word checks if the plugin works on Word.
                                            //  If you try to run the add-in in another Office application, Word-specific operations will not be started.

    document.getElementById("sideload-msg").style.display = "none"; //  Hides a “sideload” message shown during the loading phase of the add-in.
                                                                    //  In HTML, turns off the visibility of a DOM element with id “sideload-msg” (display: none).
    document.getElementById("app-body").style.display = "block";  //  The add-in makes visible an HTML element (id: app-body) representing the user interface.
                                                                  //  This DOM element is hidden by default (display: none), 
                                                                  //  but becomes visible (display: block) when the add-in is successfully installed.
    document.getElementById("customSaveButton").onclick = customSave; //  When the button is clicked, the customSave function is called.
    document.getElementById("inspectCustomPropertiesButton").onclick = inspectCustomProperties; // When the button is clicked, the inspectCustomProperties function is called.
  }
});
/*
function customSave() {
  Word.run(function (context) {
    // Get the document body
    var body = context.document.body;
    debugger;
    // Get the first paragraph to check if "Name" is already inserted
    //var firstParagraph = body.paragraphs.getFirst();
    const firstParagraph = body.paragraphs.getFirst();
    firstParagraph.load("text");

    return context.sync().then(function () {
      // Check if "Name" is at the top
      if (!firstParagraph.text.startsWith("Burak")) {
        // Insert "Name" at the top of the document
        body.insertText("Burak\n", Word.InsertLocation.start);
      }

      // Save the document
      /*return context.sync().then(function () {
        return context.document.save();
      });*//*
      return context.document.save();
    })
    //.then(context.sync)
    .then(function () {
      // Display a success message
      showNotification("Document saved with 'Burak' at the top.");
    });
  })
  .catch(function (error) {
    console.error("Error:", error);
    showNotification("An error occurred while saving the document.");
  });
}
*/

function customSave() {
  const dialogUrl = "https://localhost:3000/dialog.html?v=17";

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 40, width: 60, displayInIframe: true }, // Set size of the dialog (percentage of screen) and set the dialog to open in an iframe
    function(asyncResult){
      if(asyncResult.status === Office.AsyncResultStatus.Failed){
        console.error("Error opening dialog:", asyncResult.error.message);
        showNotification("Failed to open the dialog.");
      }else{
        const dialog = asyncResult.value;

        // Listen for messages from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(message){
          console.log("Message received from dialog:", message.message);

          if (message.message === "closeDialogAndSave") {
            dialog.close(); // Close the dialog
          }

          Word.run(function (context) {
            // Get the document body
            const body = context.document.body;
            debugger;
            // Get the first paragraph to check if "Name" is already inserted
            //var firstParagraph = body.paragraphs.getFirst();
            const firstParagraph = body.paragraphs.getFirst();
            firstParagraph.load("text");
        
            return context.sync().then(function () {
              // Check if "Name" is at the top
              if (!firstParagraph.text.startsWith("Burak")) {
                // Insert "Name" at the top of the document
                body.insertText("Burak\n", Word.InsertLocation.start);
              }
        
              // Save the document
              /*return context.sync().then(function () {
                return context.document.save();
              });*/
              return context.document.save();
            })
            //.then(context.sync)
            .then(function () {
              // Display a success message
              showNotification("Document saved with 'Burak' at the top.");
            });
          })
          .catch(function (error) {
            console.error("Error:", error);
            showNotification("An error occurred while saving the document.");
          });
        });

        // Handle dialog closed event
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          console.log("Dialog was closed.");
        });
      }
    }
  );
}

function showNotification(message) {
  const notification = document.getElementById("notification-message");
  notification.innerText = message;

  // Clear the message after 5 seconds
  setTimeout(() => {
    notification.innerText = "";
  }, 5000);
}

//  referenced by: https://stackoverflow.com/questions/44788002/word-add-in-how-to-read-custom-document-property 
function inspectCustomProperties() {  //  This function reads the Custom Properties object in the Word document and prints each property to the console.
  Word.run(function (context) { //  Initializes a set of operations in the Office.js API. Creates a context object that provides access to the Word document and makes it possible to operate on it.
                                //  In Office.js, access to a Word document is only possible within Word.run. This allows to synchronize operations between Word and the add-in.

    // Get custom properties object
    const customProperties = context.document.properties.customProperties;  //Dokümanın Custom Properties nesnesine erişir.

    // Load the custom properties
    customProperties.load("items"); //customProperties nesnesinin items özelliğini yükler.
    //Office.js API'sinde, nesneler ve özellikleri varsayılan olarak belleğe yüklenmez. İlgili verilerin kullanılabilmesi için açıkça load çağrısı yapılmalıdır.

    return context.sync().then(function () {  //  Synchronizes the load operation and retrieves the requested data from the Word document.
      //  The load operation is just a “load request”. The data is not retrieved from Word and cannot be used unless context.sync is called.
      //  If you try to access the items property without calling sync(), you may get errors like PropertyNotLoaded.
      //  The .then() method uses JavaScript's Promise construct. 
      //  Since Office.js operations are asynchronous (like context.sync), .then() is used to wait for this operation to complete and to define what to do when it completes.
      //  When a Promise completes (for example, when context.sync() runs successfully), the function specified in .then() is executed.
      //  If an error occurs during processing, .catch() is called.
      //  function () {} is a callback function.
      //  This function is called automatically when context.sync() completes.
      //  Inside this function, you write the actions you want to perform after the synchronization is complete.


      if (customProperties.items.length === 0) {  //  Check to notify the user when Custom Properties is not available.
        console.log("No custom properties found.");
      }else{
        console.log(`Found ${customProperties.items.length} custom properties:`);
        //  Iterate through the custom properties
        customProperties.items.forEach(function (property) {  //  It loops each custom property and prints the key and value pairs to the console.
          console.log(`Name: ${property.key}, Value: ${property.value}`);
        });
      }
      showNotification("Custom properties retrieved successfully. Check the console.");
    });
  })
  .catch(function (error) { //  If any error occurs, it captures the error and prints it to the console.
    console.error("Error:", error);
    showNotification("An error occurred while retrieving custom properties.");
  });
}

// Add a custom property
// key: name of custom property
// value: the value assign to the property
function addCustomProperty(key, value) {
  Word.run(function (context) {
    const customProperties = context.document.properties.customProperties;
    customProperties.add(key, value); // add new property

    return context.sync().then(function () {  // allow changes to be aplied to the Word document
      console.log(`Added custom property: { Key: ${key}, Value: ${value} }`);
      showNotification("Custom property added successfully.");
    });
  }).catch(function (error) {
    console.error("Error adding custom property:", error);
    showNotification("An error occurred while adding the custom property.");
  });
}

// Delete a custom property
// key: name of custom property
function deleteCustomProperty(key) {
  Word.run(function (context) {
    const customProperties = context.document.properties.customProperties;
    customProperties.load("items");

    return context.sync().then(function () {
      const propertyToDelete = customProperties.items.find(p => p.key === key);
      // Searches for the property in the collection whose key is equal to the given key.
      // p.key === key: The key value of each property is compared to the key parameter given to the function.
      // propertyToDelete: Represents the custom property to delete. Returns undefined if not found.

      if (propertyToDelete) {
        propertyToDelete.delete();  // delete the property
        return context.sync().then(function () {
          console.log(`Deleted custom property: ${key}`);
          showNotification("Custom property deleted successfully.");
        });
      } else {
        console.log(`Custom property not found: ${key}`);
        showNotification("Custom property not found.");
      }
    });
  }).catch(function (error) {
    console.error("Error deleting custom property:", error);
    showNotification("An error occurred while deleting the custom property.");
  });
}

// Update a custom property (delete and re-add with a new value)
function updateCustomProperty(key, value) {
  Word.run(function (context) {
    const customProperties = context.document.properties.customProperties;
    customProperties.load("items");

    return context.sync().then(function () {
      const propertyToUpdate = customProperties.items.find(p => p.key === key);
      // Searches for the property in the collection whose key is equal to the given key.
      // p.key === key: The key value of each property is compared to the key parameter given to the function.
      // propertyToUpdate: Represents the custom property to update. Returns undefined if not found.

      if (propertyToUpdate) {
        propertyToUpdate.delete(); // delete the property
        customProperties.add(key, value); // update the property
        return context.sync().then(function () {
          console.log(`Updated custom property: { Key: ${key}, Value: ${value} }`);
          showNotification("Custom property updated successfully.");
        });
      } else {
        console.log(`Custom property not found for updating: ${key}`);
        showNotification("Custom property not found for updating.");
      }
    });
  }).catch(function (error) {
    console.error("Error updating custom property:", error);
    showNotification("An error occurred while updating the custom property.");
  });
}

// Retrieve the full document text
/*
function getFullDocumentText() {
  Word.run(function (context) {
    const body = context.document.body;
    body.load("text"); // Loads the entire text of the document body (the text property) into the binding. This allows the body text to be used later.

    return context.sync().then(function () {
      console.log("Full document text:");
      console.log(body.text); // Prints all text in the document body to the console.
      showNotification("Full document text retrieved. Check the console.");
    });
  }).catch(function (error) {
    console.error("Error retrieving document text:", error);
    showNotification("An error occurred while retrieving the document text.");
  });
}
*/

// Retrieve the full document text
function getFullDocumentText() {
  Word.run(function (context) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items"); // Load all paragraphs in the document

    return context.sync().then(function () {
      // Combine paragraphs with line breaks to simulate a TXT format
      const txtFormat = paragraphs.items.map(p => p.text).join("\n");

      console.log("Full document text in TXT format:");
      console.log(txtFormat); // Print the text with line breaks preserved
      showNotification("Full document text retrieved in TXT format. Check the console.");
    });
  }).catch(function (error) {
    console.error("Error retrieving document text:", error);
    showNotification("An error occurred while retrieving the document text.");
  });
}

//  This function works similarly to the inspectCustomProperties function above, but includes some additional operations. 
//  This function is assigned to the window object, giving it global access and can be called by the developer from the console.
window.debugCustomProperties = function () {  //  Adds the function to the window object. So it can be called as debugCustomProperties() from the console.
                                              //  Starts the process to retrieve custom properties from a Word document.
                                              //  window: Represents the open browser tab (or window).
                                              //  window: Includes globally defined variables, functions and other properties.
                                              //  We want to call the debugCustomProperties function from anywhere, like the developer console, not just from taskpane.js.
  Word.run(function (context) {
    // Get the custom properties object
    const customProperties = context.document.properties.customProperties;  //  Accesses the Custom Properties object of the document.

    // Load the custom properties
    customProperties.load("items"); //  loads the items property of the customProperties object.

    // Sync the context to populate the items
    return context.sync().then(function () {  //  Synchronizes the load operation and retrieves the requested data from the Word document.
      console.log("Custom properties loaded.");
      if (customProperties.items.length === 0) {  //  Check to notify the user when Custom Properties is not available.
        console.log("No custom properties found.");
      } else {
        console.log(`Found ${customProperties.items.length} custom properties.`);
        customProperties.items.forEach(function (property) {  //  It loops each custom property and prints the key and value pairs to the console.
          console.log(`Key: ${property.key}, Value: ${property.value}`);
        });
      }
    });
  })
  .catch(function (error) { //  If any error occurs, it captures the error and prints it to the console.
    console.error("Error while inspecting custom properties:", error);
  });
};

function pushToCustomProperties(key, value) {
  Word.run(function (context) {
      const customProperties = context  .document.properties.customProperties;
      customProperties.load("items"); // Mevcut öğeleri yükler

      return context.sync().then(function () {
          // Yeni öğeyi m__items dizisine ekle
          customProperties.m__items.push({ key: key, value: value });

          console.log(`Pushed to m__items: { Key: ${key}, Value: ${value} }`);
          console.log("Updated m__items:", customProperties.m__items);
      });
  }).catch(function (error) {
      console.error("Error pushing to m__items:", error);
  });
}


function popFromCustomProperties() {
  Word.run(function (context) {
      const customProperties = context.document.properties.customProperties;
      customProperties.load("items"); // Mevcut öğeleri yükler

      return context.sync().then(function () {
          if (customProperties.m__items.length === 0) {
              console.log("No items to pop in m__items.");
              return;
          }

          // Son öğeyi kaldır
          const removedItem = customProperties.m__items.pop();

          console.log("Popped from m__items:", removedItem);
          console.log("Updated m__items:", customProperties.m__items);
      });
  }).catch(function (error) {
      console.error("Error popping from m__items:", error);
  });
}
