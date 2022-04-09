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
        //arr.push(f.name);
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
        }
        
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

        main {
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
    </header>

    <main>`;

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
