# WebPost #

WebPost is a command line tool that allows you to send HTTP requests from rainmeter. You can download it [here](http://code.google.com/p/rainmeter-outlook/downloads/detail?name=WebPost.exe) ([source](http://code.google.com/p/rainmeter-outlook/source/browse/trunk/WebPost/WebPost/mWebPost.vb)).



## Usage ##

  1. Copy WebPost.exe into your rainmeter directory.
  1. Add `WEBPOST=#PROGRAMPATH#\WebPost.exe` to the `[Variables]` section of your skin
  1. Use `!Execute ["#WEBPOST#" "-get" "http://127.0.0.1/Go"]` to send a GET-request to the specified address.