const { merge } = require('webpack-merge');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const common = require('./webpack.config.js');

module.exports = merge(common, {
    mode : 'production',
    module : {
        rules : [
            {
                test : /\.js$/,
                exclude : '/node_modules/',
                use : [
                    {
                        loader : 'babel-loader',
                        options : {
                            presets : ['@babel/preset-env']
                        }
                    }
                ]
            }
        ]
    },
    plugins : [
        new CleanWebpackPlugin()
    ]
});