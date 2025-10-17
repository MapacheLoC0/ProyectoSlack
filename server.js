require('dotenv').config({ path: 'canal.env' });
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { WebClient } = require('@slack/web-api');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());
app.use(express.static('public')); // Cambia 'index' por 'public'

const upload = multer({ dest: 'uploads/' });

// ==========================================
// FUNCIÓN: PARSEAR EXCEL AUTOMÁTICO
// ==========================================
function parseExcelAuto(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Extraer configuración (primeras 4-5 filas)
    const config = {};
    let emailStartRow = 0;

    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const key = rows[i][0]?.toString().trim().toUpperCase();
      const value = rows[i][1]?.toString().trim();

      if (key === 'NOMBRE_CANAL' || key === 'NOMBRE') {
        config.channelName = value;
      } else if (key === 'DESCRIPCION' || key === 'DESCRIPTION') {
        config.channelDescription = value;
      } else if (key === 'SLACK_TOKEN' || key === 'TOKEN') {
        config.slackToken = value;
      } else if (key === 'WORKSPACE_INVITE' || key === 'INVITE_LINK') {
        config.workspaceInviteLink = value;
      } else if (key === 'CORREO' || key === 'EMAIL' || key === 'EMAILS') {
        emailStartRow = i + 1;
        break;
      }
    }

    // Extraer correos
    const emails = [];
    for (let i = emailStartRow; i < rows.length; i++) {
      const email = rows[i][0]?.toString().trim();
      if (email && email.includes('@')) {
        emails.push(email);
      }
    }

    return { config, emails };
  } catch (error) {
    throw new Error('Error leyendo Excel: ' + error.message);
  }
}

// ==========================================
// FUNCIÓN: VALIDAR EMAIL
// ==========================================
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// ==========================================
// ENDPOINT: CREAR CANAL AUTOMÁTICO
// ==========================================
app.post('/api/auto-create-channel', upload.single('file'), async (req, res) => {
  console.log('\n🚀 ===== PROCESO AUTOMÁTICO INICIADO =====');
  
  let slackClient = null;
  let emailTransporter = null;

  try {
    if (!req.file) {
      throw new Error('No se recibió ningún archivo');
    }

    // 1. PARSEAR EXCEL
    console.log('📄 Leyendo archivo Excel...');
    const { config, emails } = parseExcelAuto(req.file.path);

    // 2. VALIDAR CONFIGURACIÓN
    if (!config.channelName) {
      throw new Error('❌ Falta NOMBRE_CANAL en el Excel (fila 1)');
    }
    if (!config.slackToken) {
      throw new Error('❌ Falta SLACK_TOKEN en el Excel (fila 3)');
    }
    if (emails.length === 0) {
      throw new Error('❌ No se encontraron correos en el Excel');
    }

    console.log(`✅ Configuración leída:`);
    console.log(`   Canal: ${config.channelName}`);
    console.log(`   Descripción: ${config.channelDescription || 'Sin descripción'}`);
    console.log(`   Token: ${config.slackToken.substring(0, 10)}...`);
    console.log(`   Correos: ${emails.length}`);

    // 3. INICIALIZAR SLACK CLIENT CON TOKEN DEL EXCEL
    slackClient = new WebClient(config.slackToken);

    // 4. VERIFICAR TOKEN
    try {
      const authTest = await slackClient.auth.test();
      console.log(`✅ Token válido - Workspace: ${authTest.team}`);
    } catch (error) {
      throw new Error('❌ Token de Slack inválido o sin permisos');
    }

    // 5. SEPARAR USUARIOS EXISTENTES VS NUEVOS
    const existingUsers = [];
    const newUsers = [];
    const invalidEmails = [];

    console.log('\n👥 Verificando usuarios en Slack...');
    for (const email of emails) {
      if (!isValidEmail(email)) {
        invalidEmails.push({ email, reason: 'Email inválido' });
        continue;
      }

      try {
        const user = await slackClient.users.lookupByEmail({ email });
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

    console.log(`\n📊 Resumen:`);
    console.log(`   ✅ En Slack: ${existingUsers.length}`);
    console.log(`   🆕 Nuevos: ${newUsers.length}`);
    console.log(`   ❌ Inválidos: ${invalidEmails.length}`);

    // 6. CREAR CANAL
    console.log(`\n🔨 Creando canal "${config.channelName}"...`);
    let channel;
    try {
      channel = await slackClient.conversations.create({
        name: config.channelName.toLowerCase().replace(/\s+/g, '-'),
        is_private: false
      });
      console.log(`   ✅ Canal creado: ${channel.channel.id}`);
    } catch (error) {
      if (error.data?.error === 'name_taken') {
        throw new Error('❌ Ya existe un canal con ese nombre');
      }
      throw new Error('❌ Error creando canal: ' + error.data?.error);
    }

    // 7. ESTABLECER DESCRIPCIÓN
    if (config.channelDescription) {
      await slackClient.conversations.setTopic({
        channel: channel.channel.id,
        topic: config.channelDescription
      });
    }

    // 8. INVITAR USUARIOS EXISTENTES
    console.log('\n👤 Invitando usuarios existentes...');
    const inviteResults = [];
    
    for (const user of existingUsers) {
      try {
        await slackClient.conversations.invite({
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
        console.log(`   ❌ Error: ${user.email} - ${error.data?.error}`);
      }
      await new Promise(resolve => setTimeout(resolve, 150));
    }

    // 9. ENVIAR EMAILS A NUEVOS USUARIOS
    let emailsSent = 0;
    if (newUsers.length > 0 && config.workspaceInviteLink) {
      console.log('\n📧 Enviando invitaciones por email...');
      
      // Configurar transporter si hay credenciales
      if (process.env.EMAIL_USER && process.env.EMAIL_PASSWORD) {
        emailTransporter = nodemailer.createTransport({
          service: 'gmail',
          auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASSWORD
          }
        });

        for (const user of newUsers) {
          try {
            await emailTransporter.sendMail({
              from: process.env.EMAIL_USER,
              to: user.email,
              subject: `Invitación a Slack - ${config.channelName}`,
              html: `
                <h2>¡Bienvenido al equipo!</h2>
                <p>Has sido invitado al canal <strong>${config.channelName}</strong> en Slack.</p>
                <p><strong>Paso 1:</strong> Únete al workspace:</p>
                <p><a href="${config.workspaceInviteLink}" style="background: #611f69; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Unirme a Slack</a></p>
                <p><strong>Paso 2:</strong> Busca el canal <strong>#${config.channelName}</strong></p>
                ${config.channelDescription ? `<p><em>${config.channelDescription}</em></p>` : ''}
              `
            });
            emailsSent++;
            console.log(`   ✅ Email enviado: ${user.email}`);
          } catch (error) {
            console.log(`   ❌ Error email: ${user.email}`);
          }
        }
      } else {
        console.log('   ⚠️  No hay credenciales de email configuradas');
      }
    }

    // 10. GUARDAR CSV DE PENDIENTES
    if (newUsers.length > 0) {
      const csvContent = 'email\n' + newUsers.map(u => u.email).join('\n');
      const csvPath = path.join(__dirname, 'uploads', `pendientes_${Date.now()}.csv`);
      fs.writeFileSync(csvPath, csvContent);
      console.log(`\n💾 CSV de pendientes: ${csvPath}`);
    }

    // 11. LIMPIAR ARCHIVO TEMPORAL
    fs.unlinkSync(req.file.path);

    // 12. RESPUESTA
    res.json({
      success: true,
      channelCreated: true,
      channelId: channel.channel.id,
      channelName: channel.channel.name,
      channelDescription: config.channelDescription || '',
      summary: {
        totalUsers: emails.length,
        existingInvited: inviteResults.filter(r => r.status === 'invited').length,
        existingFailed: inviteResults.filter(r => r.status === 'failed').length,
        newUsersEmailed: emailsSent,
        newUsersPending: newUsers.length - emailsSent,
        invalidEmails: invalidEmails.length
      },
      existingUsers: inviteResults,
      newUsers: newUsers,
      invalidEmails: invalidEmails,
      message: `✅ Canal creado exitosamente. ${inviteResults.filter(r => r.status === 'invited').length} usuarios agregados.`
    });

    console.log('\n✅ ===== PROCESO COMPLETADO =====\n');

  } catch (error) {
    console.error('\n❌ Error:', error.message);
    if (req.file?.path) {
      try { fs.unlinkSync(req.file.path); } catch {}
    }
    res.status(500).json({ 
      success: false, 
      error: error.message 
    });
  }
});

// ==========================================
// ENDPOINT: DESCARGAR PLANTILLA EXCEL
// ==========================================
app.get('/api/download-template', (req, res) => {
  try {
    // Crear libro de Excel
    const workbook = XLSX.utils.book_new();
    
    // Crear datos de la plantilla
    const data = [
      ['NOMBRE_CANAL', 'semillero-2025'],
      ['DESCRIPCION', 'Canal para el programa de formación Semillero'],
      ['SLACK_TOKEN', 'xoxb-tu-token-de-slack-aqui'],
      ['WORKSPACE_INVITE', 'https://join.slack.com/t/tu-workspace/shared_invite/xxxxx'],
      ['CORREO', ''],
      ['ejemplo1@correo.com', ''],
      ['ejemplo2@correo.com', ''],
      ['ejemplo3@correo.com', '']
    ];
    
    // Crear hoja de trabajo
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Ajustar ancho de columnas
    worksheet['!cols'] = [
      { wch: 20 },  // Columna A
      { wch: 50 }   // Columna B
    ];
    
    // Agregar hoja al libro
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Plantilla Slack');
    
    // Generar buffer
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    // Enviar archivo
    res.setHeader('Content-Disposition', 'attachment; filename=plantilla_slack.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
    
    console.log('📥 Plantilla descargada');
  } catch (error) {
    console.error('Error generando plantilla:', error);
    res.status(500).json({ error: 'Error generando plantilla' });
  }
});

// ==========================================
// ENDPOINT: VERIFICAR SALUD DEL SISTEMA
// ==========================================
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    emailConfigured: !!(process.env.EMAIL_USER && process.env.EMAIL_PASSWORD),
    timestamp: new Date().toISOString()
  });
});

// Crear carpetas necesarias
if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
if (!fs.existsSync('public')) fs.mkdirSync('public');

// Iniciar servidor
app.listen(PORT, () => {
  console.log('\n✨ ========================================');
  console.log(`   🌱 Sistema Automático SLACK`);
  console.log(`   🚀 http://localhost:${PORT}`);
  console.log(`   📧 Email: ${process.env.EMAIL_USER ? '✅' : '⚠️  No configurado'}`);
  console.log('========================================\n');
  console.log('📋 Formato del Excel requerido:');
  console.log('   Fila 1: NOMBRE_CANAL | nombre-del-canal');
  console.log('   Fila 2: DESCRIPCION | Descripción');
  console.log('   Fila 3: SLACK_TOKEN | xoxb-...');
  console.log('   Fila 4: WORKSPACE_INVITE | https://...');
  console.log('   Fila 5: CORREO');
  console.log('   Fila 6+: Un correo por fila\n');
});