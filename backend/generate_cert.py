# generate_cert.py
    from OpenSSL import crypto
    import os

    # Tạo cặp khóa
    pkey = crypto.PKey()
    pkey.generate_key(crypto.TYPE_RSA, 2048)

    # Tạo chứng chỉ tự ký
    cert = crypto.X509()
    cert.get_subject().C = "VN"
    cert.get_subject().ST = "Quang Ninh"
    cert.get_subject().L = "Cam Pha"
    cert.get_subject().O = "Cam Pha Hospital"
    cert.get_subject().OU = "IT Department"
    cert.get_subject().CN = "localhost" # Hoặc địa chỉ IP của server nếu muốn
    
    cert.set_serial_number(1000)
    cert.gmtime_adj_notBefore(0)
    cert.gmtime_adj_notAfter(10*365*24*60*60) # Hết hạn sau 10 năm

    cert.set_issuer(cert.get_subject())
    cert.set_pubkey(pkey)
    cert.sign(pkey, 'sha256')

    # Lưu file
    CERT_FILE = "cert.pem"
    KEY_FILE = "key.pem"

    with open(os.path.join(os.path.dirname(__file__), CERT_FILE), "wb") as f:
        f.write(crypto.dump_certificate(crypto.FILETYPE_PEM, cert))
    with open(os.path.join(os.path.dirname(__file__), KEY_FILE), "wb") as f:
        f.write(crypto.dump_privatekey(crypto.FILETYPE_PEM, pkey))

    print(f"Đã tạo thành công {CERT_FILE} và {KEY_FILE}")