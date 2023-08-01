function ScreenRefresh() {
        
    let wscript = new ActiveXObject('Wscript.shell');
    wscript.SendKeys('{F11}');
    }
    setInterval(myTimer, 50);
    function myTimer () {
    let CurrentDate = new Date();
    document.getElementById('curr').innerHTML = CurrentDate;
    }
    setTimeout(function()
    {
    location = ''
    },60000)