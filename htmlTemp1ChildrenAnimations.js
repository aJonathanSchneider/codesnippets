/**
 this animation will sequentally fade in all children of a DOM element
 How to: With your browser's devTools, select an html element with children, store it as a global variable 'temp1', copy&paste this code to the console and hit enter. 
 You can adjust the time with 'elemTiming' ==> in milliseconds
**/

[...temp1.children].forEach((child,idx) => { 
let elemTiming = 60;
child.animate([
  // keyframes
  { opacity: 0, transform: 'translateY(0px)' }, 
  { opacity: 1, transform: 'translateY(-150px)' }
], { 
  // timing options
  duration: elemTiming,
  delay: elemTiming*idx,
  iterations: 1,
  fill: "both"
})})


/**
 this animation will smoothly scroll through an element's children, giving each element focus for 3 seconds
**/

[...temp1.children].forEach((child,idx) => {setTimeout(() => {child.scrollIntoView({behavior: "smooth", block: "start", inline: "nearest"})}, idx*3000) })
