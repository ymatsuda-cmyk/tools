import Dexie from 'dexie'

export const db = new Dexie('gemma-chat')
db.version(1).stores({
  conversations: '++id, updatedAt',
  messages: '++id, convId, createdAt',
})

export async function createConversation(title = '新規チャット') {
  const now = Date.now()
  return db.conversations.add({ title, createdAt: now, updatedAt: now })
}

export async function touchConversation(id, title) {
  const patch = { updatedAt: Date.now() }
  if (title) patch.title = title
  await db.conversations.update(id, patch)
}

export async function deleteConversation(id) {
  await db.transaction('rw', db.conversations, db.messages, async () => {
    await db.messages.where('convId').equals(id).delete()
    await db.conversations.delete(id)
  })
}

export function listConversations() {
  return db.conversations.orderBy('updatedAt').reverse().toArray()
}

export function listMessages(convId) {
  return db.messages.where('convId').equals(convId).sortBy('createdAt')
}
