#!/bin/bash

openssl genrsa -out private.pem 2048x

openssl req -x509 -new -batch -subj "/commonName=localhost"  -key private.pem -out  public.pem -days "365"

cp public.pem public.cer

