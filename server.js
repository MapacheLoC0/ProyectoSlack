require('dotenv').config({ path: 'canal.env' });
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { WebClient } = require('@slack/web-api');
const XLSX = require('xlsx');
const Papa = require('papaparse');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer'); // npm install nodemailer

const app = express();
const PORT = process.env.PORT || 3001;

// Configuración
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const upload = multer({ dest: 'uploads/' });
const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

// Configuración de email (configura tus credenciales)
const emailTransporter = nodemailer.createTransport({
  service: 'gmail', // o tu servicio de email
  auth: {
    user: process.env.EMAIL_USER, // tu email
    pass: process.env.EMAIL_PASSWORD // tu contraseña o app password
  }
});

console.log('🔧 Configuración Sistema Híbrido Semillero:');
console.log(`   Proyecto: Semillero`);
console.log(`   App Slack: Creador`);
console.log(`   Puerto: ${PORT}`);
console.log(`   Bot Token: ${process.env.SLACK_BOT_TOKEN ? '✅' : '❌'}`);

// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

function parseFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  
  try {
    if (ext === '.csv') {
      const fileContent = fs.readFileSync(filePath, 'utf8');
      const result = Papa.parse(fileContent, {
        header: true,
        skipEmptyLines: true,
        transformHeader: header => header.trim().toLowerCase()
      });
      return result.data;
    } else {
      const workbook = XLSX.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet);
    }
  } catch (error) {
    throw new Error('Error parseando archivo: ' + error.message);
  }
}

function extractEmail(row) {
  const possibleKeys = ['email', 'correo', 'e-mail', 'mail'];
  for (const key of possibleKeys) {
    if (row[key]) return row[key].toString().trim();
  }
  return Object.values(row)[0]?.toString().trim() || null;
}

function isValidEmail(email) {
  if (!email) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// ==========================================
// ENDPOINT PRINCIPAL: SISTEMA HÍBRIDO
// ==========================================

app.post('/api/create-channel-hybrid', upload.single('file'), async (req, res) => {
  console.log('\n🚀 ===== PROCESO HÍBRIDO INICIADO =====');
  
  try {
    const { channelName, channelDescription, workspaceInviteLink } = req.body;
    
    if (!channelName || !req.file) {
      throw new Error('Faltan datos requeridos');
    }

    // Parsear archivo
    const data = parseFile(req.file.path);
    console.log(`👥 Total usuarios en archivo: ${data.length}`);

    // Separar usuarios existentes vs nuevos
    const existingUsers = [];
    const newUsers = [];
    const invalidEmails = [];

    console.log('\n📋 Verificando usuarios en Slack...');

    for (const row of data) {
      const email = extractEmail(row);
      
      if (!isValidEmail(email)) {
        invalidEmails.push({ email, reason: 'Email inválido' });
        continue;
      }

      try {
        const user = await slack.users.lookupByEmail({ email });
        existingUsers.push({ 
          email, 
          userId: user.user.id,
          name: user.user.real_name || user.user.name 
        });
        console.log(`   ✅ Existe: ${email}`);
      } catch (error) {
        if (error.data?.error === 'users_not_found') {
          newUsers.push({ email });
          console.log(`   🆕 Nuevo: ${email}`);
        } else {
          invalidEmails.push({ email, reason: error.data?.error });
        }
      }
      
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    console.log('\n📊 Resumen de usuarios:');
    console.log(`   ✅ En Slack: ${existingUsers.length}`);
    console.log(`   🆕 Nuevos: ${newUsers.length}`);
    console.log(`   ❌ Inválidos: ${invalidEmails.length}`);

    // Crear canal
    console.log(`\n🔨 Creando canal "${channelName}"...`);
    const channel = await slack.conversations.create({
      name: channelName.toLowerCase().replace(/\s+/g, '-'),
      is_private: false
    });
    console.log(`   ✅ Canal creado: ${channel.channel.id}`);

    // Establecer descripción
    if (channelDescription) {
      await slack.conversations.setTopic({
        channel: channel.channel.id,
        topic: channelDescription
      });
    }

    // Invitar usuarios EXISTENTES al canal
    console.log('\n👤 Invitando usuarios existentes...');
    const inviteResults = [];
    
    for (const user of existingUsers) {
      try {
        await slack.conversations.invite({
          channel: channel.channel.id,
          users: user.userId
        });
        inviteResults.push({ ...user, status: 'invited' });
        console.log(`   ✅ Invitado: ${user.email}`);
      } catch (error) {
        inviteResults.push({ 
          ...user, 
          status: 'failed', 
          reason: error.data?.error 
        });
        console.log(`   ❌ Error: ${user.email}`);
      }
      await new Promise(resolve => setTimeout(resolve, 150));
    }

    // Enviar invitaciones por email a usuarios NUEVOS
    let emailsSent = 0;
    if (newUsers.length > 0 && workspaceInviteLink) {
      console.log('\n📧 Enviando invitaciones por email...');
      
      for (const user of newUsers) {
        try {
          await emailTransporter.sendMail({
            from: process.env.EMAIL_USER,
            to: user.email,
            subject: `Invitación a Slack - ${channelName}`,
            html: `
              <h2>¡Bienvenido al equipo!</h2>
              <p>Has sido invitado a unirte a nuestro workspace de Slack y al canal <strong>${channelName}</strong>.</p>
              <p><strong>Paso 1:</strong> Únete al workspace haciendo clic aquí:</p>
              <p><a href="${workspaceInviteLink}" style="background: #611f69; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Unirme a Slack</a></p>
              <p><strong>Paso 2:</strong> Una vez dentro, busca el canal <strong>#${channelName}</strong></p>
              <p>¡Nos vemos pronto!</p>
            `
          });
          emailsSent++;
          console.log(`   ✅ Email enviado: ${user.email}`);
        } catch (error) {
          console.log(`   ❌ Error email: ${user.email}`);
        }
      }
    }

    // Guardar CSV con usuarios pendientes
    if (newUsers.length > 0) {
      const csvContent = 'email\n' + newUsers.map(u => u.email).join('\n');
      const csvPath = path.join(__dirname, 'uploads', `pendientes_${Date.now()}.csv`);
      fs.writeFileSync(csvPath, csvContent);
      console.log(`\n💾 CSV de pendientes guardado: ${csvPath}`);
    }

    // Limpiar archivo temporal
    fs.unlinkSync(req.file.path);

    // Respuesta completa
    res.json({
      success: true,
      channelCreated: true,
      channelId: channel.channel.id,
      channelName: channel.channel.name,
      summary: {
        totalUsers: data.length,
        existingInvited: inviteResults.filter(r => r.status === 'invited').length,
        existingFailed: inviteResults.filter(r => r.status === 'failed').length,
        newUsersEmailed: emailsSent,
        newUsersPending: newUsers.length - emailsSent,
        invalidEmails: invalidEmails.length
      },
      existingUsers: inviteResults,
      newUsers: newUsers,
      invalidEmails: invalidEmails,
      message: newUsers.length > 0 
        ? `Canal creado. ${inviteResults.length} usuarios agregados. ${newUsers.length} invitaciones pendientes por email.`
        : `Canal creado. ${inviteResults.length} usuarios agregados exitosamente.`
    });

    console.log('\n✅ ===== PROCESO COMPLETADO =====\n');

  } catch (error) {
    console.error('\n❌ Error:', error.message);
    if (req.file?.path) fs.unlinkSync(req.file.path);
    res.status(500).json({ 
      success: false, 
      error: error.message 
    });
  }
});

// ==========================================
// ENDPOINT: Verificar usuarios nuevos periódicamente
// ==========================================

app.post('/api/check-new-users', async (req, res) => {
  try {
    const { channelId, pendingEmails } = req.body;
    
    console.log(`\n🔍 Verificando ${pendingEmails.length} usuarios pendientes...`);
    
    const nowInSlack = [];
    const stillPending = [];
    
    for (const email of pendingEmails) {
      try {
        const user = await slack.users.lookupByEmail({ email });
        
        // Usuario ya está en Slack, invitarlo al canal
        await slack.conversations.invite({
          channel: channelId,
          users: user.user.id
        });
        
        nowInSlack.push({ email, status: 'added' });
        console.log(`   ✅ Usuario agregado: ${email}`);
      } catch (error) {
        if (error.data?.error === 'users_not_found') {
          stillPending.push(email);
        }
      }
      
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    
    res.json({
      success: true,
      added: nowInSlack,
      stillPending: stillPending,
      message: `${nowInSlack.length} usuarios agregados. ${stillPending.length} aún pendientes.`
    });
    
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ==========================================
// ENDPOINT: Obtener enlace de invitación del workspace
// ==========================================

app.get('/api/workspace-info', async (req, res) => {
  try {
    const authTest = await slack.auth.test();
    
    res.json({
      success: true,
      team: authTest.team,
      teamId: authTest.team_id,
      note: 'Para obtener el enlace de invitación, ve a: Configuración de Slack > Invitar personas > Generar enlace'
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Crear carpetas necesarias
if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// Iniciar servidor
app.listen(PORT, () => {
  console.log('\n✨ ========================================');
  console.log(`   🌱 Sistema Híbrido SEMILLERO`);
  console.log(`   🤖 App Slack: Creador`);
  console.log(`   🚀 http://localhost:${PORT}`);
  console.log(`   🔑 Bot Token: ${process.env.SLACK_BOT_TOKEN ? '✅' : '❌'}`);
  console.log(`   📧 Email: ${process.env.EMAIL_USER ? '✅' : '❌'}`);
  console.log('========================================\n');
  
  if (!process.env.EMAIL_USER) {
    console.log('⚠️  Para usar el sistema híbrido completo, configura en canal.env:');
    console.log('   EMAIL_USER=tu_email@gmail.com');
    console.log('   EMAIL_PASSWORD=tu_contraseña_app\n');
  }
});

app.post('/api/create-slack-channel', upload.single('file'), async (req, res) => {
  console.log('\n🚀 ===== PROCESO SLACK + EMAIL INICIADO =====');
  try {
    const { channelName, channelDescription, workspaceInviteLink } = req.body;

    if (!channelName || !req.file) {
      throw new Error('Faltan datos requeridos');
    }

    // Parsear archivo
    const data = parseFile(req.file.path);
    console.log(`👥 Total usuarios en archivo: ${data.length}`);

    // Separar usuarios existentes vs nuevos
    const existingUsers = [];
    const newUsers = [];
    const invalidEmails = [];

    for (const row of data) {
      const email = extractEmail(row);

      if (!isValidEmail(email)) {
        invalidEmails.push({ email, reason: 'Email inválido' });
        continue;
      }

      try {
        const user = await slack.users.lookupByEmail({ email });
        existingUsers.push({ 
          email, 
          userId: user.user.id,
          name: user.user.real_name || user.user.name 
        });
        console.log(`   ✅ Existe: ${email}`);
      } catch (error) {
        if (error.data?.error === 'users_not_found') {
          newUsers.push({ email });
          console.log(`   🆕 Nuevo: ${email}`);
        } else {
          invalidEmails.push({ email, reason: error.data?.error });
        }
      }

      await new Promise(resolve => setTimeout(resolve, 100));
    }

    // Crear canal
    console.log(`\n🔨 Creando canal "${channelName}"...`);
    const channel = await slack.conversations.create({
      name: channelName.toLowerCase().replace(/\s+/g, '-'),
      is_private: false
    });
    console.log(`   ✅ Canal creado: ${channel.channel.id}`);

    // Establecer descripción
    if (channelDescription) {
      await slack.conversations.setTopic({
        channel: channel.channel.id,
        topic: channelDescription
      });
    }

    // Invitar usuarios EXISTENTES al canal
    console.log('\n👤 Invitando usuarios existentes...');
    let successfulInvites = 0;
    let failedInvites = 0;
    const members = [];

    for (const user of existingUsers) {
      try {
        await slack.conversations.invite({
          channel: channel.channel.id,
          users: user.userId
        });
        successfulInvites++;
        members.push({ email: user.email, status: 'success' });
        console.log(`   ✅ Invitado: ${user.email}`);
      } catch (error) {
        failedInvites++;
        members.push({ email: user.email, status: 'failed', reason: error.data?.error });
        console.log(`   ❌ Error: ${user.email}`);
      }
      await new Promise(resolve => setTimeout(resolve, 150));
    }

    // Enviar invitaciones por email a usuarios NUEVOS
    let emailsSent = 0;
    if (newUsers.length > 0 && workspaceInviteLink) {
      console.log('\n📧 Enviando invitaciones por email...');
      for (const user of newUsers) {
        try {
          await emailTransporter.sendMail({
            from: process.env.EMAIL_USER,
            to: user.email,
            subject: `Invitación a Slack - Canal ${channelName}`,
            text: `Hola,\n\nHas sido invitado al canal "${channelName}" en Slack.\nÚnete usando este enlace: ${workspaceInviteLink}\n\nDescripción: ${channelDescription || 'Sin descripción'}\n\nSaludos,\nEquipo Semillero`
          });
          emailsSent++;
          members.push({ email: user.email, status: 'email_sent' });
          console.log(`   📧 Email enviado: ${user.email}`);
        } catch (error) {
          members.push({ email: user.email, status: 'email_failed', reason: error.message });
          console.log(`   ❌ Email error: ${user.email}`);
        }
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    // Guardar CSV con usuarios pendientes
    if (newUsers.length > 0) {
      const csvContent = 'email\n' + newUsers.map(u => u.email).join('\n');
      const csvPath = path.join(__dirname, 'uploads', `pendientes_${Date.now()}.csv`);
      fs.writeFileSync(csvPath, csvContent);
      console.log(`\n💾 CSV de pendientes guardado: ${csvPath}`);
    }

    // Limpiar archivo temporal
    fs.unlinkSync(req.file.path);

    // Respuesta
    res.json({
      success: true,
      channelCreated: true,
      channelId: channel.channel.id,
      channelName: channel.channel.name,
      invitationsSent: existingUsers.length + newUsers.length,
      successfulInvites,
      failedInvites,
      emailsSent,
      members,
      invalidEmails,
      message: `Canal creado. ${successfulInvites} usuarios agregados exitosamente. ${emailsSent} invitaciones enviadas por email.`
    });

    console.log('\n✅ ===== PROCESO SLACK + EMAIL COMPLETADO =====\n');

  } catch (error) {
    console.error('\n❌ Error:', error.message);
    if (req.file?.path) fs.unlinkSync(req.file.path);
    res.status(500).json({ 
      success: false, 
      error: error.message 
    });
  }
});