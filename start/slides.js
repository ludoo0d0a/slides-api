const google = require('googleapis');
const slides = google.slides('v1');
const drive = google.drive('v3');
const openurl = require('openurl');
const commaNumber = require('comma-number');
const Promise = require('bluebird');

// const ID_TITLE_SLIDE = 'id_title_slide';
const ID_TITLE_SLIDE = 'new_slide';


const SLIDE_TITLE_TEXT = 'Generated slide 2020';
const OUTPUT_DIR = 'generation';
const PRESENTATION_ID = '1zJLhRitVDHvjd2rjUiOxconaYV6jqKz6UFmI-_SRxGs' // template Collab

const PRESENTATION_ID_READ = '1wSc1lk8vxzsbP-a7GIYSXZJ9qxJhtyNoyw5un_PLl_s' // generated version with 2 slides

// copy from here
const ID_SOURCE_SLIDE = 'gacbe3db941_0_547';
const SLIDES_ALREADY_THERE = 2; // slides already present in slide

// true: use a master layout template ; false = copy slide from ID_SOURCE_SLIDE
const USE_TEMPLATE = false;

// true: use a predefined template
const USE_DEFAULT_TEMPLATE = false;

// true: use Promise.mapSeries to iterate batch on each slide
const USE_SEQUENTIAL_BATCH = false;

// Replace placeholders on each slide creation 
const REPLACE_ON_SLIDE = true;

const LOG_OUT = false;
const LOG_IN = false;
const LIST_INFO = false;

/**
 * Prints the number of slides and elements in a sample presentation:
 * https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
const listSlides = (r) => { 

  const [auth, ghData, presentation] = r;

  console.log('List slides ...');
  const slides = google.slides({version: 'v1', auth});

  return new Promise((resolve, reject) => {
    
    const slides = google.slides({version: 'v1', auth});
    slides.presentations.get({
      presentationId: presentation.id,
    }, (err, pres) => {
      if (err) return console.log('The API returned an error: ' + err);
      
      
      if (LIST_INFO) _listSlides(pres)
      //_listLayouts(pres)

      resolve([auth, ghData, presentation, pres.slides]);
    });

  });
}
module.exports.listSlides=listSlides;


function _listLayouts(pres){
  const layoutslength = pres.layouts.length;
  console.log('+ %s layouts:', layoutslength);
  pres.layouts.map((layout, i) => {
    console.log(`- Layout #${i + 1} "${layout.layoutProperties.displayName}"  ${layout.objectId}.`);
    _listPageElements(layout.pageElements);
  });
  console.log(' ')
}

function _listSlides(pres){
  const length = pres.slides.length;
  console.log('+ %s slides:', length);
  pres.slides.map((slide, i) => {
    console.log(`- Slide #${i + 1} ${slide.objectId}  contains ${slide.pageElements && slide.pageElements.length || 0} elements.`);
    console.log(' layout #'+slide.slideProperties.layoutObjectId+' master #'+slide.slideProperties.masterObjectId);
    console.log(' notesPage: #'+slide.slideProperties.notesPage.objectId+' '+slide.slideProperties.notesPage.pageType);
    _listPageElements(slide.pageElements);
  });
  console.log(' ');
}

function _listPageElements(pageElements){
  if (!pageElements) return;

  console.log(`>> ${pageElements && pageElements.length || 0} elements.`);
  
  pageElements
  .filter(p => p.image)
  .map((pageElement, p) => {
    if (pageElement.image){
      console.log(`  - image #${p + 1} url: ${pageElement.image.contentUrl} `);
    }else{
      console.log(`  - pageElement #${p + 1} ${pageElement.objectId}  contains ${pageElement.shape && pageElement.shape.shapeType} `);
    }

  })
  console.log(' ')
}


function _listImages(slide){
  const pageElements = slide.pageElements
  
  if (!pageElements) return;
  console.log(`>> ${pageElements && pageElements.length || 0} elements.`);
  
  const images = [];
  pageElements
  .filter(p => p.image)
  .map((pageElement, p) => {
    let image = null;
    if (pageElement.image){
      console.log(`  - image #${p + 1} url: ${pageElement.image.contentUrl} `);
      image = {
        objectId: pageElement.objectId,
        slide: slide.objectId,
        url: pageElement.image.contentUrl
      };
    }
    images.push(image)
  })
  console.log(' ')

  return images;
}


const selectFile = (auth) => { 
  const r = [
    auth,
    ghData = {},
    presentation = {
      id : PRESENTATION_ID_READ
    }
  ]
  return Promise.resolve(r);
}
module.exports.selectFile=selectFile;


const replaceTextSlides = (r) => { 

  console.log('replaceText ...');
  const [auth, ghData, presentation, _slides] = r; 

  return new Promise((resolve, reject) => {

    const allUpdateSlides = ghData.map((data, index) => updateSlideJSON(data, index, _slides));
    const slideRequests = [].concat.apply([], allUpdateSlides); // flatten the slide requests
   
    if (LOG_IN) console.log('slideRequests:', JSON.stringify(slideRequests, null, 4));

    // Execute the requests
    slides.presentations.batchUpdate({
      auth: auth,
      presentationId: presentation.id,
      resource: {
        requests: slideRequests
      }
    }, (err, res) => {
      if (err) return reject(err);
     
      if (LOG_OUT) console.log('batchUpdate response:', JSON.stringify(res, null, 4));
      resolve(r);
    });

  });
}
module.exports.replaceTextSlides=replaceTextSlides;



const replaceImages = (r) => { 
  console.log('replaceImages ...');
  const [auth, ghData, presentation, _slides] = r; 

  return new Promise((resolve, reject) => {
    const allUpdateSlides = _slides.map((slide, index) => updateImageJSON(slide, index, ghData));
    const slideRequests = [].concat.apply([], allUpdateSlides); // flatten the slide requests
   
    console.log('replaceImages slideRequests:', JSON.stringify(slideRequests, null, 4));

    if (slideRequests.length==0) return resolve(r);

    // Execute the requests
    slides.presentations.batchUpdate({
      auth: auth,
      presentationId: presentation.id,
      resource: {
        requests: slideRequests
      }
    }, (err, res) => {
      if (err) return reject(err);
     
      if (LOG_OUT) console.log('batchUpdate response:', JSON.stringify(res, null, 4));
      resolve(r);
    });

  });
}
module.exports.replaceImages = replaceImages;


function updateImageJSON(slide, index, ghData) {
  let request = [];
  const images = _listImages(slide);
  const offset = SLIDES_ALREADY_THERE ;
  // Replace image per slide
  images.forEach(image => {
    const collab = (index>=offset && ghData[index-offset])

    const newUrl = collab && collab.fields.photo;
    const imageElementId = image.objectId;
    if (imageElementId && newUrl){
      request.push({
        replaceImage: {
          imageObjectId: imageElementId,
          imageReplaceMethod: "CENTER_CROP",
          url: newUrl
        }        
        // updateImageProperties: {
        //   objectId: imageElementId,
        //   imageProperties: {
        //     link: {
        //       url: newUrl
        //     }
        //   }
        // }
      })
    }
  });

  return request;
}

const listLayouts = (r) => { 

  console.log('List layouts ...');

  const [auth, ghData, presentation] = r;
  const slides = google.slides({version: 'v1', auth});

  return new Promise((resolve, reject) => {
    
    slides.presentations.get({
      presentationId: presentation.id,
    }, (err, pres) => {
      if (err) return reject(err);

      var _layouts = pres.layouts;
      const length = pres.layouts.length;
      console.log('The presentation contains %s layouts', length);
      var layouts = {};
      for (i in _layouts) {
        layouts[_layouts[i].layoutProperties.displayName] = _layouts[i].objectId;
      }
      return resolve([auth, ghData, presentation, layouts]);
    });

  });
}

// This one with 'slideLayoutReference.predefinedLayout' works...
function createSlideJSON_default(collabData, index, slideLayout) {
  // Then update the slides.
  
  const ID_TITLE_SLIDE_TITLE = 'id_title_slide_title';
  const ID_TITLE_SLIDE_BODY = 'id_title_slide_body';
  const slideId = `${ID_TITLE_SLIDE}_${index}`;

  const request = [{
    // Creates a "TITLE_AND_BODY" slide with objectId references
    createSlide: {
      objectId: slideId,
      slideLayoutReference: {
        predefinedLayout: 'TITLE_AND_BODY'
      },
      placeholderIdMappings: [{
        layoutPlaceholder: {
          type: 'TITLE'
        },
        objectId: `${ID_TITLE_SLIDE_TITLE}_${index}`
      }, {
        layoutPlaceholder: {
          type: 'BODY'
        },
        objectId: `${ID_TITLE_SLIDE_BODY}_${index}`
      }]
    }
  }, {
    // Inserts the title
    insertText: {
      objectId: `${ID_TITLE_SLIDE_TITLE}_${index}`,
      text: `#${index + 1} {{code}}`
    }
  }, {
    // Inserts the body
    insertText: {
      objectId: `${ID_TITLE_SLIDE_BODY}_${index}`,
      text: `#${index + 1} {{name}}  {{lastname}}`
    }
  },{
    // Formats the slide paragraph's font
    updateParagraphStyle: {
      objectId: `${ID_TITLE_SLIDE_BODY}_${index}`,
      fields: '*',
      style: {
        lineSpacing: 100.0,
        spaceAbove: {magnitude: 0, unit: 'PT'},
        spaceBelow: {magnitude: 0, unit: 'PT'},
      }
    }
  }, {
    // Formats the slide text style
    updateTextStyle: {
      objectId: `${ID_TITLE_SLIDE_BODY}_${index}`,
      style: {
        bold: true,
        italic: true,
        fontSize: {
          magnitude: 10,
          unit: 'PT'
        }
      },
      fields: '*',
    }
  }];

  if (REPLACE_ON_SLIDE){
    // Replace text per slide
    for( const [key, value] of Object.entries(collabData.fields)){
      request.push({
        replaceAllText: {
          replaceText: ''+value,
          containsText: { text: '{{'+key+'}}' }
          ,pageObjectIds: [slideId]
        }
      })
    }
  }

  return request;
}

function createSlideJSON_copy(collabData, index) {
  const slideId = `${ID_TITLE_SLIDE}_${index}`;
  const originalPageId = ID_SOURCE_SLIDE;

  const lastIndex = SLIDES_ALREADY_THERE + index +1 ;

  let request = [{
    duplicateObject: {
      objectId: originalPageId,
      objectIds: {
        [originalPageId]: slideId  // => copy slide
      }
    }
  },{
    updateSlidesPosition: {
      slideObjectIds: [
        slideId
      ],
      insertionIndex: lastIndex
    }
  }];


  if (REPLACE_ON_SLIDE){
    //Replace text per slide
    for( const [key, value] of Object.entries(collabData.fields)){
      request.push({
        replaceAllText: {
          replaceText: ''+value,
          containsText: { text: '{{'+key+'}}' }
          ,pageObjectIds: [slideId]   // => TODO
        }
      })
    }
  }

  // return {request, slideId};
  return request;
}



// This one with 'slideLayoutReference.layoutId' dont work...
function createSlideJSON_custom(collabData, index, slideLayout) {
  const slideId = `${ID_TITLE_SLIDE}_${index}`;

  let request = [{
    createSlide: {
      objectId: slideId,
      slideLayoutReference: {
        layoutId: slideLayout  // => cannot replaceAllText after
      }
    }
  }];
  if (REPLACE_ON_SLIDE){
    //Replace text per slide
    for( const [key, value] of Object.entries(collabData.fields)){
      request.push({
        replaceAllText: {
          replaceText: ''+value,
          containsText: { text: '{{'+key+'}}' }
          ,pageObjectIds: [slideId]   // => failed HERE !!!!
        }
      })
    }
  }

  // return {request, slideId};
  return request;
}

function updateSlideJSON(collabData, index, slides) {

  let request = [];

  //const slideId = slides[index+1].objectId;
  const slideId = `${ID_TITLE_SLIDE}_${index}`;

  // Replace text per slide
  for( const [key, value] of Object.entries(collabData.fields)){
    request.push({
      replaceAllText: {
        replaceText: ''+value,
        containsText: { text: '{{'+key+'}}' },
        pageObjectIds: [slideId]
      }
    })
  }

    // Replace global
    request.push({
      replaceAllText: {
        replaceText: 'UPDATED ['+index+']',
        containsText: { text: '{{TITLE2}}' }
        //, pageObjectIds: [slideId]
      }
    })

    request.push({
      replaceAllText: {
        replaceText: 'UPDAT3D ['+index+']',
        containsText: { text: '{{TITLE3}}' }
        //, pageObjectIds: ['gac8e0ea071_5_74']
        , pageObjectIds: [slideId]
      }
    })


/*
request.push({
  "replaceAllShapesWithImage": {
    "imageUrl": url,
    "replaceMethod": "CENTER_INSIDE",
    "containsText": {
        "text": "{{photo}}",
    }
  }
})
*/

return request;

}

const copyFile = (authAndGHData) => { 
  const [auth, ghData] = authAndGHData;
  return new Promise((resolve, reject) => {
    // First copy the template slide from drive.


    drive.files.copy({
      auth: auth,
      // parents: [OUTPUT_DIR],
      fileId: PRESENTATION_ID,
      fields: 'id,name,webViewLink',
      resource: {
        name: SLIDE_TITLE_TEXT
      }
    }, (err, presentation) => {
      if (err) return reject(err);

      resolve([auth, ghData, presentation]);
    });
  });
}
module.exports.copyFile=copyFile;

// module.exports.updateSlides = function(r) {
//   return listSlides(r)
//   .then(replaceTextSlides);
// }

function catchPromise(r){
  return new Promise((resolve, reject) => {
    console.log('catchPromise...')
    console.log(r)
    resolve(r);
  });
}
module.exports.catchPromise=catchPromise;

//module.exports.createSlides = (authAndGHData) => new Promise((resolve, reject) => {
module.exports.createSlides = function(authAndGHData) {
  console.log('createSlides...');
  const [auth, ghData] = authAndGHData;

  const _createSlides = USE_SEQUENTIAL_BATCH ? createSlidesSeq: createSlidesAll;

  // First copy the template slide from drive.
  return copyFile(authAndGHData)
  .then(listLayouts)
  .then(_createSlides)
  //.then(catchPromise)
  ;
}

function createSlidesSeq(r) {
  const [auth, ghData, presentation, layouts] = r; 
  console.log('createSlidesSeq...');
  const slideLayout = layouts['Collab'];

  return Promise.mapSeries(ghData, function(data, index, arrayLength) {

    return new Promise((resolve, reject) => {
        console.log('--');
        console.log('--Slide' +index);
        const allSlides = createSlideJSON_custom(data, index, slideLayout);
        const slideRequests = [].concat.apply([], allSlides); // flatten the slide requests

        const slideId = `${ID_TITLE_SLIDE}_${index}`;
        //const slideId = 'gac8e0ea071_5_74';

        // // Replace global
        // slideRequests.push({
        //   replaceAllText: {
        //     replaceText: SLIDE_TITLE_TEXT,
        //     containsText: { text: '{{TITLE}}' }
        //     //,pageObjectIds: [slideId]
        //   }
        // })

        // // Replace global
        // slideRequests.push({
        //   replaceAllText: {
        //     replaceText: 'Slide '+index,
        //     containsText: { text: '{{TITLE2}}' }
        //     //,pageObjectIds: [slideId]
        //   }
        // })

        if (LOG_IN) console.log('slideRequests:', JSON.stringify(slideRequests, null, 4));
    
        // Execute the requests
        slides.presentations.batchUpdate({
          auth: auth,
          presentationId: presentation.id,
          resource: {
            requests: slideRequests
          }
        }, (err, res) => {
          if (err) {
            console.error(err.stack);
            return reject(err);
          }
          if (LOG_OUT) console.log('batchUpdate response:', JSON.stringify(res, null, 4));
          console.log('--Close slide' +index);
          resolve(res);
        });
    });

  }).then(function(result) {
    // This will run after the last step is done
    console.log("Done!")
    console.log(result); // ["1.txt", "2.txt", "3.txt", "4.txt", "5.txt"]

    return Promise.resolve([auth, ghData, presentation]);

});

}

function createSlidesAll(r) {
  const [auth, ghData, presentation, layouts] = r; 

  const slideLayout = layouts['Collab'];
  console.log('Found the layout slide "Collab"');

  return new Promise((resolve, reject) => {
      
      // default template : OK
      //const allSlides = ghData.map((data, index) => createSlideJSON_default(data, index, slideLayout));
      // Custom template : KO !!
      //const allSlides = ghData.map((data, index) => createSlideJSON(data, index, slideLayout));

      const _createSlideJson = USE_TEMPLATE ? (USE_DEFAULT_TEMPLATE ? createSlideJSON_default : createSlideJSON_custom) : createSlideJSON_copy;
      const allSlides = ghData.map((data, index) => _createSlideJson(data, index, slideLayout));

      slideRequests = [].concat.apply([], allSlides); // flatten the slide requests

      // Replace global
      slideRequests.push({
        replaceAllText: {
          replaceText: SLIDE_TITLE_TEXT,
          containsText: { text: '{{TITLE}}' }
        }
      })

      // Delete template slide
      slideRequests.push({
        deleteObject: {
          objectId: ID_SOURCE_SLIDE
        }
      })
    

      //if (LOG_IN) 
        console.log('slideRequests:', JSON.stringify(slideRequests, null, 4));

 
      // Execute the requests
      slides.presentations.batchUpdate({
        auth: auth,
        presentationId: presentation.id,
        resource: {
          requests: slideRequests
        }
      }, (err, res) => {
        if (err) {
          
          console.error(err.stack);
          return reject(err);
        }
        if (LOG_OUT) console.log('batchUpdate response:', JSON.stringify(res, null, 4));
        resolve([auth, ghData, presentation]);
      });

  });


}

// module.exports.debugInfo = (r) => new Promise((resolve, reject) => {
//   console.log('debugInfo...');
//   resolve(listSlides(r));

// });

/**
 * Opens a presentation in a browser.
 * @param {String} presentation The presentation object.
 */
module.exports.openSlidesInBrowser = (r) => {
  console.log('openSlidesInBrowser...');

  const [auth, ghData, presentation, _slides] = r; 

  console.log('Presentation URL:', presentation.webViewLink);
  openurl.open(presentation.webViewLink);
}