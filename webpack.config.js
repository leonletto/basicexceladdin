const bodyParser = require('body-parser');
const serveStatic = require('serve-static');
const path = require('path');
const exec = require('child_process').exec;
const devCerts = require('office-addin-dev-certs');
const {CleanWebpackPlugin} = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');
const WriteFilePlugin = require('write-file-webpack-plugin');
const Agent = require('agentkeepalive');
const https = require('follow-redirects').https;
const keepAliveAgent = new Agent({
  keepAlive: true,
  maxSockets: 40,
  maxFreeSockets: 10,
  timeout: 600000, // active socket keepalive for 600 seconds
  freeSocketTimeout: 300000, // free socket keepalive for 300 seconds
});

const queryBuilder = (queries1) => {
  let params = '';
  for (const query in queries1) {
    params = params + query + '=' + encodeURIComponent(queries1[query]) + '&';
  }
  if (params.length > 1) {
    params = params.substring(0, params.length - 1);
  }
  return '?' + params;
};

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const config = {
    devtool: 'source-map',
    entry: {
      taskpane: [
        'react-hot-loader/patch',
        './taskpane/index.tsx',
      ],
      commands: './commands/commands.ts'
    },
    context: path.resolve('./src'),
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js']
    },
    output: {
      path: path.resolve('dist'),
      publicPath: './',
      filename: '[name].[hash].js',
      chunkFilename: '[id].[hash].chunk.js',
    },
    optimization: {
      noEmitOnErrors: true,
      splitChunks: {
        chunks: 'async',
        minChunks: Infinity,
        name: 'vendor',
      },
    },
    module: {
      rules: [
        {
          test: /\.ts?x?$/,
          use: ['react-hot-loader/webpack', 'ts-loader'],
          include: /src/,
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          // include: /src/,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          use: {
            loader: 'file-loader',
            query: {
              name: 'assets/[name].[ext]'
            }
          }
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin([
        {
          to: 'taskpane.css',
          from: './taskpane/taskpane.css'
        }
      ]),
      new ExtractTextPlugin('[name].[hash].css'),
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './taskpane/taskpane.html',
        chunks: ['taskpane', 'vendor', 'polyfills']
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './commands/commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin([
        {
          from: '../assets',
          ignore: ['*.scss'],
          to: 'assets',
        }
      ]),
      new webpack.ProvidePlugin({
        Promise: ['es6-promise', 'Promise']
      }),
      new webpack.HotModuleReplacementPlugin(),
      {
        apply: (compiler) => {
          compiler.hooks.shouldEmit.tap('shouldEmitPlugin', () => {
            exec('rimraf dist', (err, stdout, stderr) => {
              if (stdout) process.stdout.write(stdout);
              if (stderr) process.stderr.write(stderr);
            });
          });
        },
      },
      new WriteFilePlugin(),
    ],
    devServer: {
      publicPath: '/app',
      contentBase: path.resolve('dist'),
      watchContentBase: true,
      hot: true,
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      overlay: {
        warnings: true,
        errors: true,
      },
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
      before: function (app) {
        app.use(bodyParser.json());
        app.use(
          serveStatic('src/website', {
            index: ['index.html'],
          })
        );
        app.get('/newsapitopheadlines', async (req, res0) => {
          const {} = req.headers;
          let start = Date.now();
          let ip =
            (req.headers['x-forwarded-for'] || '')
              .split(',')
              .pop()
              .trim() ||
            req.connection.remoteAddress ||
            (req.connection && req.connection.remoteAddress) ||
            req.socket.remoteAddress ||
            (req.socket && req.socket.remoteAddress) ||
            req.connection.socket.remoteAddress ||
            undefined;
          
          let options = {
            keepAliveAgent: keepAliveAgent,
            method: 'GET',
            hostname: 'newsapi.org',
            path: '/v2/top-headlines' + queryBuilder(req.query),
            headers: {
              Cookie: '__cfduid=d175352b9f486fc76e38fe1d14f2ec03d1597877974'
            }
          };
          console.log(options);
          https
            .request(options, function (res) {
              let chunks = [];

              res.on('data', function (chunk) {
                chunks.push(chunk);
              });

              res.on('end', function (chunk) {
                var body = Buffer.concat(chunks);
                let restime = Date.now() - start;
                console.log(ip + ' ' + options.hostname + options.path + ' ResponseTime:' + restime + 'ms');
                res0.send(body);
              });

              res.on('error', function (error) {
                console.error(error);
              });


            })
            .end();
        });
        console.log('Listening on https://localhost:' + process.env.npm_package_config_dev_server_port || 3000);
      }
    }
  };

  return config;
};
