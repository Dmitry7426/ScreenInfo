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
        if (a1[i].innerHTML.match(/на месте/) || a1[i].innerHTML.match(/на смене/) || a1[i].innerHTML.match(/подработка/)) {
            var Id = a1[i].dataset.testId
            a1[i].style.color = 'green';
            a1[i-1].style.color = 'green';
	    }

        if (a1[i].innerHTML.match(/нет/) || a1[i].innerHTML.match(/больничный/) || a1[i].innerHTML.match(/обучение/) || a1[i].innerHTML.match(/командировка/) || a1[i].innerHTML.match(/отпуск/) || a1[i].innerHTML.match(/работа из дома/)) {
            var Id = a1[i].dataset.testId
            a1[i].style.color = 'red';
            a1[i-1].style.color = 'red';
            }
    
  


	if (a1[i].innerHTML.match(/Клининг/)){
            var Id = a1[i].dataset.testId
            a1[i].style.color = 'black';
            a1[i-1].style.color = 'black';
	    a1[i].style.fontWeight = 'bold';
	    a1[i].innerText = ' \n \n \n Клининг';
        }
    }
});

   