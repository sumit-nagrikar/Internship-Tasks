const express = require('express');
const cors = require('cors');
const uploadRoute = require('./routes/upload.route');
const oauthRoute = require('./routes/oauth.route')

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use('/upload', uploadRoute);


// app.use('/', oauthRoute) // just for authentication

//For creatting token.json for first time means for completing authentication with google oauth2
// const { getAuthenticatedClient } = require('./utils/googlesheet.utils');

// getAuthenticatedClient().then(() => {
//   console.log('Authentication successful!');
// });

app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});