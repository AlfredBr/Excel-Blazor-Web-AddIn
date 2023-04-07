"use strict";
window.onresize = () => {
    console.log(`window.onresize(${window.innerWidth}, ${window.innerHeight})`);
    console.assert(window.dotNetHelper);
    window.dotNetHelper?.invokeMethodAsync('OnResize', window.innerWidth, window.innerHeight);
};
function SetDotNetHelper(dotNetHelper) {
    window.dotNetHelper = dotNetHelper;
    console.assert(window.dotNetHelper);
    window.dotNetHelper?.invokeMethodAsync('OnResize', window.innerWidth, window.innerHeight);
}