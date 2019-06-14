const rimraf = require('rimraf');

rimraf.sync('./outputCSV')
rimraf.sync('./outputJSON')
rimraf.sync('./cleanJSON')
console.log('successfully removed directories');
