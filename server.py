import socket
import threading
import openpyxl
from datetime import datetime

HOST = '192.168.240.60'
PORT = 12346


class ChatServer:
    def __init__(self, host, port):
        self.host = host
        self.port = port
        self.clients = {}
        self.server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.server.bind((self.host, self.port))
        self.server.listen(5)
        self.workbook = openpyxl.Workbook()
        self.ws = self.workbook.active
        self.ws.title = "Chat logs"
        self.ws.append(["Timestamp", "Sender", "Recipient", "Message"])

    def log_message(self, sender, recipient, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.ws.append([timestamp, sender, recipient, message])
        self.workbook.save("chat_logs.xlsx")

    def start(self):
        print(f"[*] Listening on {self.host}:{self.port}")
        while True:
            client_socket, addr = self.server.accept()
            client_handler = threading.Thread(target=self.handle_client, args=(client_socket, addr))
            client_handler.start()

    def handle_client(self, client_socket, addr):
        username = client_socket.recv(1024).decode('utf-8')
        self.clients[username] = client_socket
        print(f"[*] {username} connected from {addr[0]}:{addr[1]}")

        while True:
            try:
                message = client_socket.recv(1024).decode('utf-8')
                if message:
                    if message == f"{username} has left the chat.":
                        print(message)
                        self.broadcast(message, client_socket)
                        self.remove_client(username)
                        break
                    elif message.startswith('@'):
                        recipient, msg = message.split(' ', 1)
                        recipient = recipient[1:]
                        if recipient in self.clients:
                            self.clients[recipient].send(f"[Private from {username}]: {msg}".encode('utf-8'))
                            self.log_message(username, recipient, msg)
                        else:
                            client_socket.send(f"User '{recipient}' not found.".encode('utf-8'))
                    else:
                        self.broadcast(f"[{username}]: {message}", client_socket)
                        self.log_message(username, "All", message)
                else:
                    self.remove_client(username)
                    break
            except Exception as e:
                print(f"Error handling client {username}: {e}")
                self.remove_client(username)
                break

    def broadcast(self, message, sender_socket):
        for client in self.clients.values():
            if client != sender_socket:
                try:
                    client.send(message.encode('utf-8'))
                except:
                    client.close()
                    self.remove_client(client)

    def remove_client(self, username):
        if username in self.clients:
            self.clients[username].close()
            del self.clients[username]
            print(f"{username} disconnected")


if __name__ == "__main__":
    server = ChatServer(HOST, PORT)
    server.start()
