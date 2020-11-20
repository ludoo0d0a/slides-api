const read = require('read-file');

/**
 * Get the license data from BigQuery and our license data.
 * @return {Promise} A promise to return an object of licenses keyed by name.
 */
module.exports.getCollabsData = (auth) => new Promise((resolve, reject) => {

  read('collabs.json', 'utf8', (err, buffer) => {
    if (err) return reject(err);
    const data = JSON.parse(buffer);

    resolve([auth, data]);
  });

})