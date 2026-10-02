"""
LINE Group ID Catcher
รัน script นี้เพื่อจับ Group ID จาก LINE Webhook event
"""
import http.server
import json
import sys

PORT = 8888
group_id_found = []

class WebhookHandler(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        
        try:
            data = json.loads(body.decode('utf-8'))
            print("\n=== LINE Webhook Event Received ===")
            print(json.dumps(data, indent=2, ensure_ascii=False))
            
            # Extract groupId
            for event in data.get('events', []):
                source = event.get('source', {})
                if source.get('type') == 'group':
                    gid = source.get('groupId', '')
                    if gid:
                        print(f"\n✅ LINE GROUP ID FOUND: {gid}")
                        group_id_found.append(gid)
        except Exception as e:
            print(f"Parse error: {e}")
        
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'OK')
    
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        if group_id_found:
            self.wfile.write(f"GROUP ID: {group_id_found[0]}".encode())
        else:
            self.wfile.write(b"Waiting for LINE events...")
    
    def log_message(self, format, *args):
        pass  # suppress default logs

if __name__ == '__main__':
    server = http.server.HTTPServer(('', PORT), WebhookHandler)
    print(f"🚀 Webhook server started on port {PORT}")
    print(f"   กำลังรอ LINE Webhook events...")
    print(f"   ถ้าใช้ ngrok: ngrok http {PORT}")
    print(f"   แล้วนำ URL ไปตั้ง Webhook ใน LINE Developer Console\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
        if group_id_found:
            print(f"✅ Group ID: {group_id_found[0]}")
