## pip install ##
python3 -m pip install flask
python3 -m pip install jinja2
python3 -m pip install gunicorn

## add to service ##
echo y | cp ./word2img.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable word2img
systemctl start word2img

## add config to nginx ##
echo y | cp ./word2img.nginx.http.conf /etc/nginx/conf.d/
systemctl restart nginx ## please make sure nginx installed ##
