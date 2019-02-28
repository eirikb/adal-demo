let popUp = true;

if (window.location.hash.match(/nopopup/)) {
  console.log('disable PopUp');
  popUp = false;
}

const resource = 'https://graph.microsoft.com';
const config = {
  popUp,
  clientId: '8cc739ac-fd78-4687-b9e8-13d8e461150c',
  callback(errorDesc, token, error, tokenType) {
    console.log('callback');
    callback({errorDesc, token, error, tokenType});
  }
};

const authContext = new AuthenticationContext(config);

document.querySelector('#login').addEventListener('click', () => authContext.login());
document.querySelector('#resource').addEventListener('click', getResource);
document.querySelector('#resourcepopup').addEventListener('click', getResourcePopup);
document.querySelector('#getuser').addEventListener('click', getUser);
document.querySelector('#faketokentimeout').addEventListener('click', fakeTokenTimeout);

const info = document.querySelector('#info');

function showHide(...ids) {
  document.querySelectorAll('button').forEach(el => el.style.display = 'none');
  for (let id of ids) {
    document.querySelector(`#${id}`).style.display = '';
  }
}

function callback({errorDesc, token, error, tokenType}) {
  if (errorDesc) {
    console.log(error, '::', errorDesc);
  }
  if (tokenType === 'id_token') {
    getResource();
  } else if (error === 'login required') {
    showHide('login');
  } else if (error) {
    // Final hope
    showHide('resourcepopup');
  } else if (token) {
    showHide();
    getUserWithToken(token);
  }
}

function getResource() {
  console.log('getResource');
  showHide();

  authContext.acquireToken(resource, (errorDesc, token, error) => {
    console.log('acquireToken');
    callback({errorDesc, token, error});
  });
}

function getResourcePopup() {
  console.log('getResourcePopup');
  showHide();

  authContext.acquireTokenPopup(resource, null, null, (errorDesc, token, error) => {
    console.log('acquireTokenPopup');
    callback({errorDesc, token, error});
  });
}

function getUser() {
  console.log('getUser');
  showHide();

  let token = authContext.getCachedToken(resource);
  if (token) {
    getUserWithToken(token);
  } else {
    getResource();
  }
}

function getUserWithToken(token) {
  console.log('getUserWithToken', token.slice(0, 20), '...');
  showHide();

  info.innerText = 'Loading user...';
  fetch('https://graph.microsoft.com/v1.0/me', {
    headers: {
      accept: 'application/json',
      authorization: `Bearer ${token}`
    }
  }).then(res => res.json()).then(me => {
    info.innerText = 'Done!';
    console.log('done', me);
    showHide('getuser', 'faketokentimeout');
  });
}

function fakeTokenTimeout() {
  sessionStorage['adal.expiration.keyhttps://graph.microsoft.com'] = 1;
}

if (authContext.isCallback(window.location.hash)) {
  authContext.handleWindowCallback();
  console.log('isCallback');
} else {
  console.log('init');
  showHide();
  document.querySelector('#app').style.display = 'block';
  let user = authContext.getCachedUser();
  if (user) {
    getUser();
  } else {
    showHide('login');
  }
}