/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            app.showNotification('Instruction:',"Use the arrow buttons to show the Vowels or Consonants. Select the character form the grid to print it on documents.");
            $('#get-data-from-selection').click(getDataFromSelection);
            $('.but').click(paste);
            $('#back').click(vow);
            $('#forward').click(con);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }



    function paste() {
        var data= $(this).val();
        Office.context.document.setSelectedDataAsync(data, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
        });

        // Function that writes to a div with id='message' on the page.
        function write(message) {
            document.getElementById('message').innerText += message;
        }


    }
   
    function vow() {
        document.getElementById('data').innerHTML = "Vowels";
        document.getElementById('consonents').style.visibility = "hidden";
        document.getElementById('vowel').style.visibility = "visible";




    }
    function con() {
        document.getElementById('data').innerHTML = "Consonants";
        document.getElementById('consonents').style.visibility = "visible";
        document.getElementById('vowel').style.visibility = "hidden";
        document.getElementById('vowel').style.position = "fixed";
        document.getElementById('vowel').style.display = "block";
        document.getElementById('vowel').style.width = "100%";






    }

    ;


})();