# Documentation

This is a [Node.js](https://nodejs.org/en) project powered by [Express JS Framework](https://expressjs.com/). It is implemented with [Admin Kit V3.4.0](https://adminkit.io/) which is a free [Bootstrap 5](https://getbootstrap.com) admin template.

## Getting Started

First, install all the dependencies form 'package.json', run this command:

```bash
npm install
```
Then, create a file name `.env`, `copy and paste` these configuration variables with proper values:

```bash
# Port
PORT=8081

# Mode
MODE=<mode>

# Database uri
MONGODB_URI=<mongodb_uri>

# Production url
PROD_URL=https://admin.leavenowgrow.com

# Secret key
SECRET_KEY=<secret_key>

JWT_SECRET=<jwt_secret>

JWT_EXPIRES=<jwt_expires>

DEFAULT_PASS=<default_password>

# Credentials
TENANT_ID=<tenent_id>
CLIENT_ID=<client_id>
USER_EMAIL=<user_email>

CLIENT_SECRET=<client_secret>

# Endpoints
AAD_ENDPOINT=<aad_endpoint>
GRAPH_ENDPOINT=<graph_endpoint>
```

Then, run the development server:

```bash
npm run dev
```

Or, run the production server:

```bash
npm run start
```

Open [http://localhost:8081](http://localhost:8081) with your browser to see the result.

You can start editing the app by modifying `server.js` and other files.