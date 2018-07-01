var text = document.querySelector('.text');
var svg = document.querySelector('svg');

var support = 'd' in text.style;

if (!support) {

	// 	bring text in center
	text.classList.add('center');

	// 	blur svg element
	svg.classList.add('blur');
}