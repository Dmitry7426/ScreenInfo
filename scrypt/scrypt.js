function ScreenRefresh() {
    var wscript = new ActiveXObject('Wscript.shell');
    wscript.SendKeys('{F11}');
    }
    setInterval(myTimer, 50);
    function myTimer () {
    let d = new Date();
    document.getElementById('curr').innerHTML = d;
    }
    setTimeout(function()
    {
    location = ''
    },60000)


window.addEventListener("load", function() {
    var a1 = document.getElementsByTagName('a');
    for (let i = 0; i <= a1.length; i++) {
        console.log(a1[i].innerHTML)
        if (a1[i].innerHTML.match(/на месте/)){
            var Id = a1[i].dataset.testId
            a1[i].style.color = 'green';
            a1[i-1].style.color = 'green';
	}

	if (a1[i].innerHTML.match(/Клининг/)){
            var Id = a1[i].dataset.testId
            a1[i].style.color = 'black';
            a1[i-1].style.color = 'black';
	    a1[i].style.textDecoration = 'underline';
        }
    }
});

   