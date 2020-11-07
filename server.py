from socket import socket


class Server:
	"""
		Represents the head of the server.
	"""
	
	def __init__(self):
		"""
			All the attributes and the data of the server.
		"""
		
		self.

	def main_loop(self):
		"""
			# The main loop of the server. Every loop, handle the new clients, the messages that need to be sent and the messages that received.
		"""
		
		while True:
			
			# Get the received messages and see which clients are waiting for response.
			rlist, wlist, xlist = select.select([server_socket] + open_clients_sockets, open_clients_sockets, [])
			
			# Handle all the clients that are waiting for response.
			self.send_messages(wlist)
			
			# Iterate over all the clients that sent a message.
			for current_socket in rlist:
				
				# If the current socket equals to the server socket it's trying to connect, accept it and append it to open_clients_sockets.
				if current_socket is server_socket:

					(new_socket, address) = server_socket.accept()
					
					# Receive the client's name.
					client_name = new_socket.recv(1024)
					
					# Create a random number between 10**16 - 10**20
					client_id = generate_id(occupied_ids)
					
					# Add it to the occupied ids.
					occupied_ids.append(client_id)
					
					# Create the client.
					new_client = Client(client_name, client_id, new_socket)
					
					# Append the new client instance to the connected clients list.
					connected_clients.append(new_client)
					
					# Append the client socket to open_clients_sockets
					open_clients_sockets.append(new_socket)
					
					# Send to the client its id so he can identify next time he sends data to the server, before the id o represents the client has an opponent and 1 that he has to wait.
					try:
						
						# If there is a client that currently waiting for an opponent, send the current client that an opponent was found and set the game.
						if waiting_for_opponent:

							new_socket.send('new_message:0' + client_id)
							
							set_game(new_client, waiting_for_opponent)

						else:

							new_socket.send('new_message:1' + client_id)
							
							waiting_for_opponent = new_client
							
					except:

						disconnected_clients.append(new_socket)
						continue
					
					
				# See what the current client sent.
				else:
				
					try:
						
						data = current_socket.recv(1024)
						
					except:

						disconnected_clients.append(new_socket)
						continue
					
					# This client left the server. 
					if data == "":
						
						disconnected_clients.append(current_socket)
						continue
					
					# Extract the id of the client from the data he sent.
					client_id = number_length(data)
					
					# Find the client that it's his id.
					client = get_client_by_id(client_id)
					
					# Update the data- now without the id of the client in its beginning.
					data = data[len(client_id):]

					if data[:14] == "game_property:":
						
						update_game_property(client, data[14:])

			# Iterate over all the players that are currently waiting for data.
			for current_socket in wlist:

				current_client = get_client_by_socket(current_socket)
				
				# If there are two players that are looking for opponent, set their game.
				if not current_client.opponent and waiting_for_opponent and waiting_for_opponent is not current_client:
					
					set_game(current_socket, waiting_for_opponent)
			
			# Manage everything related to the game itself.
			for game in current_games:
				
				game.game_manager()
			
			# Iterate over all the clients sockets that have been disconnected. 
			for dis_socket in disconnected_clients:
				
				# If the client that left was waiting for opponent set waiting for opponent to None.
				if waiting_for_opponent and dis_socket is waiting_for_opponent.socket:

					waiting_for_opponent = None
					
				open_clients_sockets.remove(dis_socket)
				dis_socket.shutdown(socket.SHUT_RDWR)
				dis_socket.close()
				
			disconnected_clients = []
		
	
class Client():
	"""
		A client that is currently connected to the server.
	"""
	
	def __init__(self):
		"""
			Initialize the new client attributes.
		"""
		
		pass