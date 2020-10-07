/**
 this animation will sequentally fade in all children of a DOM element
 How to: With your browser's devTools, select an html element with children, store it as a global variable 'temp1', copy&paste this code to the console and hit enter. 
 You can adjust the time with 'elemTiming' ==> in milliseconds
**/

[...temp1.children].forEach((child,idx) => { 
let elemTiming = 100;
child.animate([
  // keyframes
  { opacity: 0 }, 
  { opacity: 1 }
], { 
  // timing options
  duration: elemTiming,
  delay: elemTiming*idx,
  iterations: 1,
  fill: "both"
})})
