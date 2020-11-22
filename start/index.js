
require('dotenv').config();
const auth = require('./auth');
const slides = require('./slides');
const collabs = require('./collabs');

function generateSlide(){
  console.log('-- Start generating slides. --')
  auth.getClientSecrets()
    .then(auth.authorize)
    .then(collabs.getCollabsData)
    .then(slides.createSlides)
    .then(slides.listSlides)
    .then(slides.replaceImages)
    //.then(slides.replaceTextSlides)
    //.then(slides.debugInfo)
    //.then(slides.catchPromise)
    .then(slides.openSlidesInBrowser)
    .then(() => {
      console.log('-- Finished generating slides. --');
    });
}


function readSlide(){
  console.log('-- Start reading slides. --')
  auth.getClientSecrets()
    .then(auth.authorize)
    .then(slides.selectFile)
    .then(slides.listSlides)
    //.then(slides.debugInfo)
    //.then(slides.catchPromise)
   // .then(slides.openSlidesInBrowser)
    .then(() => {
      console.log('-- Finished reading slides. --');
    });
}


generateSlide()

//readSlide();