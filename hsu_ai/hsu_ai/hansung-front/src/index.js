import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { GoogleOAuthProvider } from '@react-oauth/google';

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  <GoogleOAuthProvider clientId="1050471225985-cf0r2c76775h8vl44vb7e4rd3q7dvmek.apps.googleusercontent.com">
    <App />
  </GoogleOAuthProvider>
);
