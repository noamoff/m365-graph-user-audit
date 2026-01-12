import 'dotenv/config'
import fetch from 'isomorphic-fetch'
import { Client } from '@microsoft/microsoft-graph-client'

async function getAccessToken() {
  const response = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    }
  )

  const data = await response.json()
  return data.access_token
}

async function main() {
  const token = await getAccessToken()

  const client = Client.init({
    authProvider: done => done(null, token)
  })

  const users = await client.api('/users').select('id,displayName,mail').get()

  console.log('Users:')
  users.value.forEach(u =>
    console.log(`- ${u.displayName} (${u.mail})`)
  )
}

main().catch(console.error)
