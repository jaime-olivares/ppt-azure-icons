Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("app-body").style.display = "block";
        document.getElementById("app-body").onclick = processClick;
    }
});

let position = 50;

function getLines(text: string, limit: number): string[]
{
    var lines = [];
    var first = true;
    var line = "";

    var parts = text.split(" ");

    for (var i=0; i<parts.length; i++)
    {
        if (first)
        {
            line = parts[i];
            first = false;
        }
        else if (line.length + parts[i].length + 1 <= limit)
        {
            line += " " + parts[i];
        }
        else
        {
            lines.push(line);
            line = parts[i];
        }
    }

    if (line.length > 0)
        lines.push(line);

    return lines;
}

async function processSvg(data: string, label: string): Promise<string>
{
    data = data.replace("</defs>", "</defs><g>").replace("</svg>", "</g></svg>");

    var parser = new DOMParser();
    var xmlDoc = parser.parseFromString(data, "text/xml"); 

    xmlDoc.documentElement.setAttribute("viewBox", "-16 0 50 40");
    xmlDoc.documentElement.setAttribute("width", "50");
    xmlDoc.documentElement.setAttribute("height", "40");

    var lines = getLines(label, 15);
    var y = 22;

    for (var i=0; i<lines.length; i++)
    {
        var node = xmlDoc.createElement("text");
        node.setAttributeNS(null, "x", "9");
        node.setAttributeNS(null, "y", y.toString());
        node.setAttributeNS(null, "width", "100%");
        node.setAttributeNS(null, "height", "auto");
        node.setAttributeNS(null, "font-size", "5");
        node.setAttributeNS(null, "font-family", "monospace");
        node.setAttributeNS(null, "text-anchor", "middle");
        node.setAttributeNS(null, "dominant-baseline", "hanging");
        node.appendChild(document.createTextNode(lines[i]));    
        xmlDoc.childNodes[0].appendChild(node);

        y += 5;
    }

    var serializer = new XMLSerializer();
    return serializer.serializeToString(xmlDoc);
}

export async function processClick() 
{
    let event = window.event;
    let combo = <HTMLSelectElement>document.getElementById("icon-size");
    let iconsize: number = parseInt(combo.options[combo.selectedIndex].text);

    const response = await fetch(event.target["src"]);

    if (response.ok)
    {
        var data = await processSvg(await response.text(), event.target["title"]);

        var options: Office.SetSelectedDataOptions = 
        { 
            coercionType: Office.CoercionType.XmlSvg, 
            imageWidth: iconsize / 18 * 50, 
            imageLeft: position, 
            imageTop: position 
        };

        // await Office.context.document.setSelectedDataAsync(data, { coercionType: Office.CoercionType.Text });

        await Office.context.document.setSelectedDataAsync(data, options);

        position += iconsize / 2;

        if (position > 450)
          position = 50;
    }
}
