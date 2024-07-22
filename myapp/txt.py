from telethon import TelegramClient, events

# Remember to use your own values from my.telegram.org!
api_id = 24889174
api_hash = '3f2d0e15946f00c6fe70788fe8945d9b'
client = TelegramClient('session_name', api_id, api_hash)

# @client.on(events.NewMessage)
# async def my_event_handler(event):
#     if 'hello' in event.raw_text:
#         await event.reply('hi!')

@client.on(events.NewMessage(chats=-1001568547659))
async def my_event_handler(event):
    print(event.raw_text)

async def main():
    # Start the client
    await client.start()

    # Print information about yourself
    me = await client.get_me()
    print(me.stringify())
    print(f'Username: {me.username}')
    print(f'Phone: {me.phone}')

    # Print all dialogs/conversations you are part of
    async for dialog in client.iter_dialogs():
        print(dialog.name, 'has ID', dialog.id)

    # Send a message to yourself
    await client.send_message('me', 'Hello, myself!')

    # Keep the script running
    await client.run_until_disconnected()

with client:
    client.loop.run_until_complete(main())


