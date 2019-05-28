const rimraf = require('rimraf');

rimraf.sync('./outputCSV')
rimraf.sync('./outputJSON')
console.log('successfully removed directories');
