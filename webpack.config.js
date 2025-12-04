const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isProduction = argv.mode === 'production';

  return {
    entry: {
      commands: './src/commands/commands.ts',
      taskpane: './src/taskpane/taskpane.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].bundle.js',
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.jsx']
    },
    module: {
      rules: [
        {
          test: /\.(ts|tsx)$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              presets: [
                '@babel/preset-react',
                '@babel/preset-typescript'
              ]
            }
          }
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'assets', to: 'assets' }
        ]
      })
    ],
    devServer: {
      static: [
        {
          directory: path.join(__dirname, 'dist')
        },
        {
          directory: path.join(__dirname, 'assets'),
          publicPath: '/assets'
        }
      ],
      port: 3000,
      hot: true,
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      https: true // Required for Office Add-ins
    },
    devtool: isProduction ? false : 'inline-source-map'
  };
};