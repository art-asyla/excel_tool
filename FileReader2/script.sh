#!/bin/sh

var="this is var from bash"
sed -n '/start/,/end/p' output.txt
echo $var

sleep 2s
export $var