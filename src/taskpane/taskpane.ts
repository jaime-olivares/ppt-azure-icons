Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "block";
        document.getElementById("app-body").onclick = processClick;
    }
});

let position = 50;

export async function processClick() 
{
    let event = window.event;
    let combo = <HTMLSelectElement>document.getElementById("icon-size");
    let iconsize: number = parseInt(combo.options[combo.selectedIndex].text);

    const response = await fetch(event.target["src"]);

    if (response.ok)
    {
        const data = await response.text();

        var options: Office.SetSelectedDataOptions = 
        { 
            coercionType: Office.CoercionType.XmlSvg, 
            imageWidth: iconsize, 
            imageLeft: position, 
            imageTop: position 
        };

        await Office.context.document.setSelectedDataAsync(data, options);

        /*
        await PowerPoint.run(async (context) => {
            var shapes = context.presentation.slides.getItemAt(0).;
            var textbox = shapes.addTextBox("Hello!", 
                { 
                  left: 100, 
                  top: 300, 
                  height: 300, 
                  width: 450 
                });
            textbox.name = "Textbox";
          
            return context.sync();
          });
*/        

        position += iconsize / 2;

        if (position > 450)
          position = 50;



    }
}

