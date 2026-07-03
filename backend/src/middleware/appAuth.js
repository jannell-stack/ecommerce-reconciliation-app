const APP_PASSWORD = process.env.APP_PASSWORD || 'P@ss12345678';

function requireAppAuth(req, res, next) {
  const password = req.headers['x-app-password'];
  if (password && password === APP_PASSWORD) {
    return next();
  }
  return res.status(401).json({ error: 'Unauthorized' });
}

module.exports = {
  APP_PASSWORD,
  requireAppAuth,
};
