const google = require('googleapis');
const slides = google.slides('v1');
const drive = google.drive('v3');
const openurl = require('openurl');
const commaNumber = require('comma-number');
const Promise = require('bluebird');

const ID_TITLE_SLIDE = 'id_title_slide';
const ID_TITLE_SLIDE_TITLE = 'id_title_slide_title';
const ID_TITLE_SLIDE_BODY = 'id_title_slide_body';

const SLIDE_TITLE_TEXT = 'CM Lux 2020';
const PRESENTATION_ID = '1zJLhRitVDHvjd2rjUiOxconaYV6jqKz6UFmI-_SRxGs'

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
      const length = pres.slides.length;
      console.log('The presentation contains %s slides:', length);
      pres.slides.map((slide, i) => {
        console.log(`- Slide #${i + 1} ${slide.objectId}  contains ${slide.pageElements && slide.pageElements.length || 0} elements.`);
      });


      const pages = pres.pages;
      //const pageslength = pres.pages.length;

      resolve([auth, ghData, presentation, pres.slides]);
    });

  });
}

const replaceTextSlides = (r) => { 

  console.log('replaceText ...');
  const [auth, ghData, presentation, _slides] = r; 

  return new Promise((resolve, reject) => {

    const allUpdateSlides = ghData.map((data, index) => updateSlideJSON(data, index, _slides));
    const slideRequests = [].concat.apply([], allUpdateSlides); // flatten the slide requests
    console.log(slideRequests);  


    // Execute the requests
    slides.presentations.batchUpdate({
      auth: auth,
      presentationId: presentation.id,
      resource: {
        requests: slideRequests
      }
    }, (err, res) => {
      if (err) {
        console.log(JSON.stringify(slideRequests, null, 4));
        reject(err);
      } else {
          resolve(r);
      }
    });

  });
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
  const ID_TITLE_SLIDE = 'id_title_slide';
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

  return request;
}

// This one with 'slideLayoutReference.layoutId' works...
function createSlideJSON(collabData, index, slideLayout) {
  // Then update the slides.
  const slideId = `slide_collab_${index}`;

  let request = [{
    createSlide: {
      objectId: slideId,
      slideLayoutReference: {
        layoutId: slideLayout
      }
    }
  }]; 

  // Replace text per slide
  for( const [key, value] of Object.entries(collabData.fields)){
    request.push({
      replaceAllText: {
        replaceText: ''+value,
        containsText: { text: '{{'+key+'}}' }
        //,pageObjectIds: [slideId]
      }
    })
  }

  return request;
}

function updateSlideJSON(collabData, index, slides) {

  let request = [];

  //const slideId = slides[index+1].objectId;
  const slideId = `${ID_TITLE_SLIDE}_${index}`;
/*
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
*/
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


module.exports.updateSlides = function(r) {
  return listSlides(r)
  .then(replaceTextSlides);
}

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

  // First copy the template slide from drive.
  return copyFile(authAndGHData)
  //.then(r => listLayouts(r))
  .then(listLayouts)
  .then(createSlidesAll)
  //.then(createSlidesSeq)
  
  //.then(catchPromise)
  ;
}

function createSlidesSeq(r) {
  const [auth, ghData, presentation, layouts] = r; 
  console.log('createSlidesSeq...');
  const slideLayout = layouts['Collab'];

  return Promise.mapSeries(ghData, function(data, index, arrayLength) {

    console.log('--');
    console.log('--Slide' +index);
    const allSlides = createSlideJSON(data, index, slideLayout);
    const slideRequests = [].concat.apply([], allSlides); // flatten the slide requests
    console.log(JSON.stringify(slideRequests, null, 4));
  
    const slideId = `${ID_TITLE_SLIDE}_${index}`;
    //const slideId = 'gac8e0ea071_5_74';

    // Replace global
    slideRequests.push({
      replaceAllText: {
        replaceText: SLIDE_TITLE_TEXT,
        containsText: { text: '{{TITLE}}' }
      }
    })

    // Replace global
    slideRequests.push({
      replaceAllText: {
        replaceText: 'Slide '+index,
        containsText: { text: '{{TITLE2}}' }
        //,pageObjectIds: [slideId]
      }
    })

    return new Promise((resolve, reject) => {
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
          console.log(JSON.stringify(res, null, 4));
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

  
  const allSlides = ghData.map((data, index) => createSlideJSON_default(data, index, slideLayout));
  //const allSlides = ghData.map((data, index) => createSlideJSON(data, index, slideLayout));
  slideRequests = [].concat.apply([], allSlides); // flatten the slide requests
  console.log(slideRequests);

  // Replace global
  slideRequests.push({
    replaceAllText: {
      replaceText: SLIDE_TITLE_TEXT,
      containsText: { text: '{{TITLE}}' }
    }
  })

  return new Promise((resolve, reject) => {
    
      // Execute the requests
      slides.presentations.batchUpdate({
        auth: auth,
        presentationId: presentation.id,
        resource: {
          requests: slideRequests
        }
      }, (err, res) => {
        if (err) return reject(err);
        console.log(JSON.stringify(res, null, 4));
        resolve([auth, ghData, presentation]);
      });

  });


}

module.exports.debugInfo = (r) => new Promise((resolve, reject) => {
  console.log('debugInfo...');
  resolve(listSlides(r));

});

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