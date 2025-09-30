<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->
<!DOCTYPE html>
<html>

<head>
    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>

<body>
    <p>This add-in will insert the text of the button at the place of your cursor</p>
    <button onclick="addTemplateTextToEmail(potscheepvaart)">Pot Scheepvaart</button><br><br>
    <button onclick="addTemplateTextToEmail(hollandshipstores)">Holland Ship Stores</button><br><br>
    <button onclick="addTemplateTextToEmail(loodsvier)">Loods 4</button><br><br>
    <button onclick="addTemplateTextToEmail(shipko)">Shipko</button><br><br>
    <button onclick="addTemplateTextToEmail(vrolijk)">Rederij Vrolijk</button><br><br>
    <button onclick="addTemplateTextToEmail(verbrugge)">Verbrugge</button><br><br>
    <button onclick="addTemplateTextToEmail(papiercontainer)">Papiercontainer</button><br><br>
    <?php
        echo "<p><strong>Hello World from PHP!</strong></p>";
    ?>
</body>

<script>
    var potscheepvaart = "<p style='font-family:\"Courier New\"'> p/a Pot Scheepvaart<br> Handelskade West 28A <br> 9934 AA Delfzijl <br> Nederland </br></p>"
    var hollandshipstores = "<p style='font-family:\"Courier New\"'> p/a Holland Ship Stores<br> Trambaan 2-4 <br> 9936 ES Farmsum <br> Nederland </br></p>"
    var loodsvier = "<p style='font-family:\"Courier New\"'> p/a Warehouse Wagenborg Shipping B.V.<br>Loods 4 <br>  Visserijweg 3<br> 9936 HB Farmsum<br> Nederland </br></p>"
    var shipko = "<p style='font-family:\"Courier New\"'> p/a Shipko Shipsupply<br> Pieter Zeemanweg 135-137<br> 3316 GZ Dordrecht<br> Nederland </br></p>"
    var vrolijk = "<p style='font-family:\"Courier New\"'> p/a Rederij Vrolijk<br> Dr. Lelykade 3H<br> 2583 CL Scheveningen<br> Nederland </br></p>"
    var verbrugge = "<p style='font-family:\"Courier New\"'> p/a Verbrugge Marine BV<br> Luxemburgweg 2<br> 4455 TM Nieuwdorp<br> Port Nr. 6700<br>Nederland </br></p>"
    var papiercontainer = "<p> Zouden jullie woensdag a.s. onze papiercontainer kunnen legen?<br><br> Order Details<br> C.V. Scheepvaartonderneming Doggersbank<br> Handelskade West 28A<br> 9934AA Delfzijl<br> the Netherlands<br> <br> Purchase Number: Doggersbank-20240119133610 <br> COC: 02332995<br> VAT Number: NL 8046.24.677.B01<br> RSIN: 804624677<br> EORI: NL804624677<br> <br> General Contact: info@pot-scheepvaart.com<br> Invoices To: invoice@pot-scheepvaart.com<br> </p>"

    /**
     * Writes the text of the variable in the email body.
     */
    function addTemplateTextToEmail(templateText) {
        Office.context.mailbox.item.body.setSelectedDataAsync(
            templateText,
            {
                coercionType: "html", // Write text as HTML
            },

            // Callback method to check that setAsync succeeded
            function (asyncResult) {
                if (asyncResult.status ==
                    Office.AsyncResultStatus.Failed) {
                    write(asyncResult.error.message);
                }
            }
        );
    }

</script>

</html>
