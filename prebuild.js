const fs = require('fs')

console.log("prebuild started");

var getTitle = function(str)
{
    var parts = str.replace(".svg", "").split("-");
    var title = parts.slice(3).join(" ");

    return title;
}

var all = {};

fs.readdirSync("./azure_icons/Icons", { withFileTypes: true}).forEach(f => {
    if (f.isDirectory && !f.name.startsWith('.'))
    {
        all[f.name] = [];
        
        fs.readdirSync("./azure_icons/Icons/" + f.name, { withFileTypes: true}).forEach(ff => {
            all[f.name].push(ff.name)
        });        
    }
});

var content = `<!-- Copyright (c) Jaime Olivares. Licensed under the MIT License. -->
<!-- This is an auto-generated file. Do not edit it manually. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Azure Icons Add-in</title>

    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <style>
        html,
        body {
            width: 100%;
            height: 100%;
            margin: 5px;
            padding: 0;
            font-family: Arial, Helvetica, sans-serif;
            font-size: 10pt;
            -webkit-touch-callout: none;
            -webkit-user-select: none;
             -khtml-user-select: none;
               -moz-user-select: none;
                -ms-user-select: none;
                    user-select: none;        }
        
        img { 
            width: 48px;
            height: 48px;
            margin: 5px;
        }
        
        header {
            position: fixed;
            width: 100%;
            top: 0px;
            height: 25px;
            padding: 5px 0px;
            background-color: white;
        }

        #info-button {
            cursor: pointer;
            margin-left: 10px;
        }

        #icons {
            padding-top: 10px;
            display: block;
        }

        #faq {
            display: none;
            padding-top: 10px;
        }
    </style>
</head>

<body id="app-body">
    <header>
        <span>Size: </span>
        <select name="iconsize" id="icon-size">
            <option value="32">32</option>
            <option value="48">48</option>
            <option value="64">64</option>
            <option value="96">96</option>
            <option value="128">128</option>
        </select>
        <span id="info-button">&#9432;<span>
    </header>

    <div id="faq">
        <h3>What file formats are available for icons?</h3>
        Our icon files are in Scalable Vector Graphic (SVG) format, so you can resize them without hurting quality.
        <h3>Do they work in PowerPoint?</h3>
        SVGs work in Office 365 versions of PowerPoint. Use them the same same way you use JPGs, PNGs and other image files.
        <h3>How can see all the icons at once?</h3>
        While not quite available yet, we plan to offer a PPT version of the icons in a future release through Azure Architecture Center.
        <h3>Are these icon assets available in Visio or as VSS files?</h3>
        Not yet, but they will be very soon. We are working with the Visio product team. In the meantime, you can use the Import function in Visio. Learn more about importing icons to Visio.
        <h3>What if I need a different size?</h3>
        You can resize these files to any size. See question 1 at the top of the page.
        <h3>Why can’t I find an icon?</h3>
        We try to be complete, but omissions happen. If you can’t find an icon, please email CnESymbols@microsoft.com.
        <h3>Can I modify these icons?</h3>
        Icons may not be cropped, flipped or rotated, and their shape may not be distorted or changed. Please visit the Azure Architecture Center where you downloaded these icons to understand the Azure Brand guidelines and legal use agreement.
        <h3>Do the icons need to be labeled?</h3>
        Yes, the full name of the service should always appear near the icon, but not so close that it looks like it’s overlapping the icon.
        <h3>Can I use these icons as part of representative images for my product or service?</h3>
        No, Azure icons may only be used to represent the Microsoft product that it was designed for.
    </div>

    <main id="icons">`;

        for (const [key, value] of Object.entries(all)) 
        {
            content += "<p>" + key + "</p>\n";

            for (var i=0; i < value.length; i++)
            {
                content += '<img src="../../azure_icons/Icons/' + key + '/' + value[i] + '" title="' + getTitle(value[i]) + '" />\n';            
            }
        }    

        content += `</main>
</body>
</html>`;

fs.writeFileSync("./src/taskpane/taskpane.html", content);
