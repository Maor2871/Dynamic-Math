import socket


class Server:
	"""
		Represents the head of the server.
	"""

	def __init__(self, ip, port):
		"""
			All the attributes and the data of the server.
		"""

		# The ip of the server.
		self.ip = ip

		# The port of the server.
		self.port =  port

		# The socket of the server.
		self.socket = socket.socket()

		# A list of all the currently connected clinets.
		self.connected_clienet = []

		# A list of all the clients that are no longer connected and need the get disconnected.
		self.disconnected_clients = []

		# All the sockets that are connected to the server.
		self.open_clients_sockets = []

	def start(self):
		"""
			The function starts the server.
		"""

		# Set the server on air.
		self.socket.bind((self.ip, self.port))

		# Start the main loop.
		self.main_loop()

	def main_loop(self):
		"""
			The main loop of the server. Every loop, handle the new clients, the messages that need to be sent and the messages that received.
		"""

		while True:

			# Get the received messages and see which clients are waiting for response.
			rlist, wlist, xlist = select.select([self.socket] + open_clients_sockets, open_clients_sockets, [])

			# Handle all the clients that are waiting for response.
			self.send_messages(wlist)

			# Iterate over all the clients that sent a message.
			for current_socket in rlist:

				# If the current socket equals to the server socket it's trying to connect, accept it and append it to open_clients_sockets.
				if current_socket is self.socket:

					(new_socket, address) = server_socket.accept()

					# Receive the client's name.
					client_name = new_socket.recv(1024)

					# Create a random number between 10**16 - 10**20
					client_id = generate_id(occupied_ids)

					# Add it to the occupied ids.
					occupied_ids.append(client_id)

					# Create the client.
					new_client = Client(client_name, new_socket)

					# Append the new client instance to the connected clients list.
					self.connected_clients.append(new_client)

					# Append the client socket to the list with all the currently connected sockets.
					self.open_clients_sockets.append(new_socket)

				# See what the current client sent.
				else:

					# Try to receive the message of the client.
					try:

						# Receive the message of the client.
						data = current_socket.recv(1024)

					# Something is wrong with the client's connection. Disconnect the client from the server.
					except:

						# Add the client's sockets to the list of sockets that need to be disconnected.
						disconnected_clients.append(current_socket)
						continue

					# The client told the server that it has disconnected.
					if data == "":

						# Add the socket of the client to the list of sockets that need to be disconnected.
						disconnected_clients.append(current_socket)
						continue

					# Get the current client by its socket.
					client = self.get_client(current_socket)

					# Deal with the client's request.
					self.handle_request(client, data)

			# Iterate over all the players that are currently waiting for data.
			for current_socket in wlist:

				# Get the client by it socket.
				current_client = self.get_client(current_socket)

				# If this client is supposed to receive a mesasge, send him his message.
				if current_client in self.messages_to_send:

					# Send the message to the client.
					current_client.send_message(self.messages_to_send[current_client])

				# If not.
				else:

					# something went wrong. Send the client an empty message indicating that the server can't handle his request.
					current_client.send_message("")

			# Iterate over all the clients sockets that have been disconnected.
			for dis_socket in self.disconnected_clients:

				open_clients_sockets.remove(dis_socket)
				dis_socket.shutdown(socket.SHUT_RDWR)
				dis_socket.close()

			# Empty the list of disconnected clients.
			self.disconnected_clients = []

	def get_client(self, current_socket):
		"""
			The function receives a client socket and finds the client of the socket from the clients list.
		"""

		# Iterate over all the clients registered to the server.
		for client in self.client:

			# Check if the current client owns the received socket.
			if client.socket is current_socket:

				# Return the client.
				return client

		# There's no client registered to the server with the received socket.
		return None

	def handle_request(self, client, message):
		"""
			The function receives a client and a message. It handles the request of the client.
		"""

		pass

	def disconnect_client(self, client):
		"""
			The function receives a client and disconnects him from the server.
		"""

		# Add the socket of the client to the disconnected clients sockets.

class Client():
	"""
		A client that is currently connected to the server.
	"""

	def __init__(self, server, socket):
		"""
			Initialize the new client attributes.
		"""

		# Reference to the server.
		self.server = server

		# The socket of the client.
		self.socket = socket

	def send_message(self, message):
		"""
			The function receives a message and sends it to the client.
		"""

		# Try to send the message to the client.
		try:

			# Send the message.
			self.socket.send(message)

		except:

			# Something is wrong with the client's connection. Disconnect him from the server.
			self.server.disconnect_client(self)

server = Server("0.0.0.0", 327)
server.start()
