var iframe = document.createElement('iframe');
iframe.style.display = 'none';
document.body.appendChild(iframe);
window.alert = iframe.contentWindow.alert;
window.prompt = iframe.contentWindow.prompt;
var Attack = Object.values(document.querySelector('#app > divide > divide'))[1].child[1]._owner.stateNode;
var lure = parseInt(prompt('Lure rank value? (1 to 5)')) - 1;
Attack.setState({
   lure: lure
});
alert(`Lure rank set to ${lure + 1}!`);
