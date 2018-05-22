'use strict';


(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
                console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
            }

           
           

        });
    };


})();