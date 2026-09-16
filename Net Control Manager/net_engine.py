import json
import asyncio
import websockets
from PyQt6.QtCore import QThread, pyqtSignal


class WebSocketHostThread(QThread):
    message_received = pyqtSignal(dict)

    def __init__(self, host="0.0.0.0", port=8765, parent=None):
        super().__init__(parent)
        self.host = host
        self.port = port
        self.loop = None
        self.server = None
        self.connected_clients = set()
        self.get_current_state_callback = None
        self.is_running = True

    def run(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        async def handler(websocket):
            self.connected_clients.add(websocket)
            # Push full snapshot to newly connected client
            if self.get_current_state_callback:
                snapshot = self.get_current_state_callback()
                if snapshot:
                    await websocket.send(json.dumps(snapshot))

            try:
                async for message in websocket:
                    try:
                        data = json.loads(message)
                        self.message_received.emit(data)
                        # Re-broadcast to all other connected peers
                        broadcast_tasks = [
                            client.send(message)
                            for client in self.connected_clients
                            if client != websocket
                        ]
                        if broadcast_tasks:
                            await asyncio.gather(*broadcast_tasks, return_exceptions=True)
                    except json.JSONDecodeError:
                        pass
            except websockets.exceptions.ConnectionClosed:
                pass
            finally:
                self.connected_clients.remove(websocket)

        async def start_server():
            self.server = await websockets.serve(handler, self.host, self.port)
            await asyncio.Future()

        try:
            self.loop.run_until_complete(start_server())
        except Exception:
            pass

    def send_data(self, data_dict):
        if self.loop and self.connected_clients:
            msg = json.dumps(data_dict)
            for client in list(self.connected_clients):
                asyncio.run_coroutine_threadsafe(client.send(msg), self.loop)

    def stop(self):
        self.is_running = False
        if self.loop:
            self.loop.call_soon_threadsafe(self.loop.stop)


class WebSocketClientThread(QThread):
    message_received = pyqtSignal(dict)
    connected = pyqtSignal()
    disconnected = pyqtSignal(str)

    def __init__(self, uri="ws://127.0.0.1:8765", parent=None):
        super().__init__(parent)
        self.uri = uri
        self.loop = None
        self.websocket = None
        self.is_running = True

    def run(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        async def connect():
            try:
                async with websockets.connect(self.uri) as ws:
                    self.websocket = ws
                    self.connected.emit()
                    async for message in ws:
                        try:
                            data = json.loads(message)
                            self.message_received.emit(data)
                        except json.JSONDecodeError:
                            pass
            except Exception as e:
                self.disconnected.emit(str(e))

        self.loop.run_until_complete(connect())

    def send_data(self, data_dict):
        if self.loop and self.websocket:
            msg = json.dumps(data_dict)
            asyncio.run_coroutine_threadsafe(self.websocket.send(msg), self.loop)

    def stop(self):
        self.is_running = False
        if self.loop:
            self.loop.call_soon_threadsafe(self.loop.stop)