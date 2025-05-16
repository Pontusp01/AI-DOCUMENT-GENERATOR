module.exports = (config) => {
  const webpack = require('webpack');
  const dotenv = require('dotenv');
  
  // Hämta miljövariabler från .env
  const env = dotenv.config().parsed || {};
  
  // Konvertera miljövariabler till format som kan användas i klientkod
  const envKeys = Object.keys(env).reduce((result, key) => {
    result[`process.env.${key}`] = JSON.stringify(env[key]);
    return result;
  }, {});
  
  // Lägg till en plugin som gör miljövariablerna tillgängliga
  config.plugins.push(new webpack.DefinePlugin(envKeys));
  
  return config;
};