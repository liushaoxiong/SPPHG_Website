var oMarquee = document.getElementById("mq"); //��������
var iLineHeight = 25; //���и߶ȣ�����
var iLineCount = 7; //ʵ������
var iScrollAmount = 1; //ÿ�ι����߶ȣ�����
function run() {
oMarquee.scrollTop += iScrollAmount;
if ( oMarquee.scrollTop == iLineCount * iLineHeight )
oMarquee.scrollTop = 0;
if ( oMarquee.scrollTop % iLineHeight == 0 ) {
window.setTimeout( "run()", 3000 );
} else {
window.setTimeout( "run()", 2 );
}
}
oMarquee.innerHTML += oMarquee.innerHTML;
window.setTimeout( "run()", 2000 );