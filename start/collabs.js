const read = require('read-file');

/**
 * Get the license data from BigQuery and our license data.
 * @return {Promise} A promise to return an object of licenses keyed by name.
 */
module.exports.getCollabsData = (auth) => new Promise((resolve, reject) => {

  read('collabs.json', 'utf8', (err, buffer) => {
    if (err) return reject(err);
    let data = JSON.parse(buffer);

    //Return list of 1 items
    //data = data.slice(0,1)

    resolve([auth, data]);
  });

})