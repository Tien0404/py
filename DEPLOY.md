# Hướng dẫn Deploy lên VPS

## 1. Chuẩn bị VPS
```bash
# Cài Python
sudo apt update
sudo apt install python3 python3-pip python3-venv -y
```

## 2. Upload code
```bash
# Tạo folder
mkdir -p /var/www/nrl-lookup
cd /var/www/nrl-lookup

# Upload các file sau:
# - app.py
# - requirements.txt
# - nrl.xlsx
# - templates/index.html
```

## 3. Cài đặt
```bash
# Tạo virtual environment
python3 -m venv venv
source venv/bin/activate

# Cài thư viện
pip install -r requirements.txt
```

## 4. Chạy với Gunicorn (production)
```bash
# Cài gunicorn
pip install gunicorn

# Chạy
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

## 5. Chạy như service (tự động khởi động)
```bash
sudo nano /etc/systemd/system/nrl-lookup.service
```

Nội dung:
```ini
[Unit]
Description=NRL Lookup Service
After=network.target

[Service]
User=www-data
WorkingDirectory=/var/www/nrl-lookup
Environment="PATH=/var/www/nrl-lookup/venv/bin"
ExecStart=/var/www/nrl-lookup/venv/bin/gunicorn -w 4 -b 0.0.0.0:5000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable nrl-lookup
sudo systemctl start nrl-lookup
```

## 6. Cấu hình Nginx (optional)
```bash
sudo nano /etc/nginx/sites-available/nrl-lookup
```

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

```bash
sudo ln -s /etc/nginx/sites-available/nrl-lookup /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

## Lỗi thường gặp

### 1. File nrl.xlsx không tồn tại
- Đảm bảo đã upload file `nrl.xlsx` lên VPS

### 2. Permission denied
```bash
sudo chown -R www-data:www-data /var/www/nrl-lookup
```

### 3. Port 5000 bị chặn
```bash
sudo ufw allow 5000
```

### 4. Xem log lỗi
```bash
sudo journalctl -u nrl-lookup -f
```
