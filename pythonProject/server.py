import socket
import threading
import openpyxl
from openpyxl import Workbook
from datetime import datetime

HOST = '192.168.0.113'
PORT = 12346

clients = {}

wb = Workbook()
ws = wb.active
ws.title = "Chat logs"

ws.append(["Timestamp", "Sender", "Recipient", "Message"])


def log_message(sender, recipient, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([timestamp, sender, recipient, message])
    wb.save("chat_logs.xlsx")


def handle_client(client_socket, addr):
    username = client_socket.recv(1024).decode('utf-8')
    clients[username] = client_socket
    print(f"[*] {username} connected from {addr[0]}:{addr[1]}")

    while True:
        try:
            message = client_socket.recv(1024).decode('utf-8')
            if message:
                if message == f"{username} has left the chat.":
                    print(message)
                    broadcast(message, client_socket)
                    remove_client(username)
                    break
                elif message.startswith('@'):
                    recipient, msg = message.split(' ', 1)
                    recipient = recipient[1:]
                    if recipient in clients:
                        clients[recipient].send(f"[Private from {username}]: {msg}".encode('utf-8'))
                        log_message(username, recipient, msg)
                    else:
                        client_socket.send(f"User '{recipient}' not found.".encode('utf-8'))
                else:
                    broadcast(f"[{username}]: {message}", client_socket)
                    log_message(username, "All", message)
            else:
                remove_client(username)
                break
        except Exception as e:
            print(f"Error handling client {username}: {e}")
            remove_client(username)
            break


def broadcast(message, sender_socket):
    for client in clients.values():
        if client != sender_socket:
            try:
                client.send(message.encode('utf-8'))
            except:
                client.close()
                remove_client(client)


def remove_client(username):
    if username in clients:
        clients[username].close()
        del clients[username]
        print(f"{username} disconnected")


def start_server():
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.bind((HOST, PORT))
    server.listen(5)
    print(f"[*] Listening on {HOST}:{PORT}")

    while True:
        client_socket, addr = server.accept()
        client_handler = threading.Thread(target=handle_client, args=(client_socket, addr))
        client_handler.start()


if __name__ == "__main__":
    start_server()
