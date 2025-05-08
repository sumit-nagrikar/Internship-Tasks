const { getAuthenticatedClient } = require('./googlesheet.utils');

getAuthenticatedClient().then(() => {
  console.log('Authentication successful!');
});
