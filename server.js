require('dotenv').config({ path: 'canal.env' });
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { WebClient } = require('@slack/web-api');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const sgMail = require('@sendgrid/mail');
const { addPending, checkPendingInvites } = require('./bot-monitor');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const upload = multer({ dest: 'uploads/' });

// ==========================================
// FUNCIÓN: PARSEAR EXCEL SIMPLIFICADO
// ==========================================
function parseExcelAuto(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const config = {};
    let emailStartRow = 0;

    // Buscar configuración en las primeras filas
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const key = rows[i][0]?.toString().trim().toUpperCase();
      const value = rows[i][1]?.toString().trim();

      if (key === 'NOMBRE_CANAL' || key === 'NOMBRE' || key === 'CANAL') {
        config.channelName = value;
      } else if (key === 'DESCRIPCION' || key === 'DESCRIPTION' || key === 'DESC') {
        config.channelDescription = value;
      } else if (key === 'CORREO' || key === 'EMAIL' || key === 'EMAILS' || key === 'CORREOS') {
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
// FUNCIÓN: VALIDAR CONFIGURACIÓN DEL SISTEMA
// ==========================================
function validateSystemConfig() {
  const missing = [];
  
  if (!process.env.SLACK_BOT_TOKEN) missing.push('SLACK_BOT_TOKEN');
  if (!process.env.WORKSPACE_INVITE_LINK) missing.push('WORKSPACE_INVITE_LINK');
  if (!process.env.SENDGRID_API_KEY) missing.push('SENDGRID_API_KEY');
  if (!process.env.EMAIL_FROM) missing.push('EMAIL_FROM');
  
  if (missing.length > 0) {
    throw new Error(`❌ Faltan variables de entorno: ${missing.join(', ')}`);
  }
}

// ==========================================
// ENDPOINT: CREAR CANAL AUTOMÁTICO
// ==========================================
app.post('/api/auto-create-channel', upload.single('file'), async (req, res) => {
  console.log('\n🚀 ===== PROCESO AUTOMÁTICO INICIADO =====');
  
  let slackClient = null;

  try {
    // 0. VALIDAR CONFIGURACIÓN DEL SISTEMA
    validateSystemConfig();
    
    if (!req.file) {
      throw new Error('No se recibió ningún archivo');
    }

    // 1. PARSEAR EXCEL
    console.log('📄 Leyendo archivo Excel...');
    const { config, emails } = parseExcelAuto(req.file.path);

    // 2. VALIDAR DATOS DEL EXCEL
    if (!config.channelName) {
      throw new Error('❌ Falta NOMBRE_CANAL en el Excel (debe estar en la primera fila)');
    }
    if (emails.length === 0) {
      throw new Error('❌ No se encontraron correos en el Excel');
    }

    console.log(`✅ Configuración leída del Excel:`);
    console.log(`   Canal: ${config.channelName}`);
    console.log(`   Descripción: ${config.channelDescription || 'Sin descripción'}`);
    console.log(`   Correos: ${emails.length}`);
    console.log(`\n🔧 Configuración del sistema (.env):`);
    console.log(`   Token: ${process.env.SLACK_BOT_TOKEN.substring(0, 15)}...`);
    console.log(`   Workspace: ${process.env.WORKSPACE_INVITE_LINK.substring(0, 40)}...`);

    // 3. INICIALIZAR SLACK CLIENT CON TOKEN DEL .ENV
    slackClient = new WebClient(process.env.SLACK_BOT_TOKEN);

    // 4. VERIFICAR TOKEN
    try {
      const authTest = await slackClient.auth.test();
      console.log(`✅ Token válido - Workspace: ${authTest.team}`);
    } catch (error) {
      throw new Error('❌ Token de Slack inválido o sin permisos (verifica SLACK_BOT_TOKEN en .env)');
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

    // 6. CREAR CANAL PRIVADO
    console.log(`\n🔨 Creando canal privado "${config.channelName}"...`);
    let channel;
    try {
      channel = await slackClient.conversations.create({
        name: config.channelName.toLowerCase().replace(/\s+/g, '-'),
        is_private: true
      });
      console.log(`   ✅ Canal privado creado: ${channel.channel.id}`);
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

    // 8. INVITAR USUARIOS EXISTENTES AL CANAL
    console.log('\n👤 Invitando usuarios existentes al canal...');
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

    // 9. ENVIAR EMAILS A NUEVOS USUARIOS CON SENDGRID
    let emailsSent = 0;
    const emailErrors = [];
    
    if (newUsers.length > 0) {
      console.log('\n📧 Enviando invitaciones por email con SendGrid...');
      
      // Configurar SendGrid
      sgMail.setApiKey(process.env.SENDGRID_API_KEY);
      
      for (const user of newUsers) {
        try {
          await sgMail.send({
            to: user.email,
            from: process.env.EMAIL_FROM,
            subject: `🔔 Invitación a Slack - Canal #${config.channelName}`,
            html: `
              <!DOCTYPE html>
              <html>
              <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
              </head>
              <body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;">
                <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
                  
                  <!-- Header -->
                  <div style="background: linear-gradient(135deg, #611f69 0%, #4a154b 100%); padding: 40px 30px; text-align: center;">
                    <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: 700;">
                      ¡Bienvenido a Slack! 👋
                    </h1>
                    <p style="color: #f8e6ff; margin: 10px 0 0 0; font-size: 16px;">
                      Has sido invitado a unirte al canal
                    </p>
                    <div style="background-color: rgba(255,255,255,0.2); display: inline-block; padding: 8px 20px; border-radius: 20px; margin-top: 15px;">
                      <span style="color: #ffffff; font-size: 18px; font-weight: 600;">#${config.channelName}</span>
                    </div>
                  </div>

                  <!-- Content -->
                  <div style="padding: 40px 30px;">
                    
                    <!-- Alert Box -->
                    <div style="background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 20px; border-radius: 4px; margin-bottom: 30px;">
                      <p style="margin: 0; color: #856404; font-size: 15px; line-height: 1.5;">
                        <strong>⚠️ Importante:</strong> Solo necesitas unirte al workspace, serás añadido automáticamente al canal privado
                      </p>
                    </div>

                    <!-- Step 1 -->
                    <div style="margin-bottom: 35px;">
                      <div style="background-color: #611f69; color: white; width: 36px; height: 36px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-weight: bold; font-size: 18px; margin-bottom: 15px;">
                        1
                      </div>
                      <h2 style="margin: 10px 0; color: #1d1c1d; font-size: 20px; font-weight: 600;">
                        Únete al workspace de Unicatólica
                      </h2>
                      <p style="color: #616061; margin: 10px 0 20px 0; font-size: 15px; line-height: 1.6;">
                        Haz clic en el botón de abajo para unirte:
                      </p>
                      <div style="text-align: center;">
                        <a href="${process.env.WORKSPACE_INVITE_LINK}" 
                           style="display: inline-block; background-color: #611f69; color: #ffffff; padding: 15px 40px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 16px; box-shadow: 0 2px 8px rgba(97, 31, 105, 0.3);">
                          🚀 Unirme al Workspace
                        </a>
                      </div>
                    </div>

                    <!-- Step 2 -->
                    <div style="margin-bottom: 35px;">
                      <div style="background-color: #28a745; color: white; width: 36px; height: 36px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-weight: bold; font-size: 18px; margin-bottom: 15px;">
                        2
                      </div>
                      <h2 style="margin: 10px 0; color: #1d1c1d; font-size: 20px; font-weight: 600;">
                        ¡Listo! Serás añadido automáticamente
                      </h2>
                      <p style="color: #616061; margin: 10px 0; font-size: 15px; line-height: 1.6;">
                        Una vez que te unas al workspace, nuestro bot te agregará automáticamente al canal privado <strong>#${config.channelName}</strong> en unos minutos.
                      </p>
                      <div style="background-color: #e8f5e9; padding: 15px; border-radius: 6px; margin-top: 15px;">
                        <p style="margin: 0; color: #2e7d32; font-size: 14px;">
                          🤖 <strong>Sistema automático:</strong> El bot revisa cada 30 minutos si te has unido al workspace y te añade al canal sin que tengas que hacer nada más.
                        </p>
                      </div>
                    </div>

                    ${config.channelDescription ? `
                    <!-- Description Box -->
                    <div style="background-color: #f8f9fa; border-left: 4px solid #007a5a; padding: 20px; border-radius: 4px; margin-bottom: 30px;">
                      <h3 style="margin: 0 0 10px 0; color: #1d1c1d; font-size: 16px; font-weight: 600;">
                        📋 Sobre este canal:
                      </h3>
                      <p style="margin: 0; color: #616061; font-size: 14px; line-height: 1.6;">
                        ${config.channelDescription}
                      </p>
                    </div>
                    ` : ''}

                    <!-- Help -->
                    <div style="margin-top: 35px; padding-top: 25px; border-top: 1px solid #e0e0e0;">
                      <p style="color: #616061; font-size: 14px; margin: 0; line-height: 1.6;">
                        ¿Necesitas ayuda? Responde este correo o contacta al administrador del workspace.
                      </p>
                    </div>

                  </div>

                  <!-- Footer -->
                  <div style="background-color: #f8f8f8; padding: 25px 30px; text-align: center; border-top: 1px solid #e0e0e0;">
                    <p style="color: #999999; font-size: 12px; margin: 0; line-height: 1.5;">
                      Este es un correo automático del Sistema de Gestión Slack<br>
                      <strong>Unicatólica</strong> - Semillero de Investigación
                    </p>
                  </div>

                </div>
              </body>
              </html>
            `
          });
          
          emailsSent++;
          console.log(`   ✅ Email enviado: ${user.email}`);

          // ✅ AGREGAR A PENDIENTES PARA EL BOT MONITOR
          addPending(user.email, channel.channel.id, channel.channel.name);
          
        } catch (error) {
          emailErrors.push({ email: user.email, reason: error.message });
          console.log(`   ❌ Error email: ${user.email} - ${error.message}`);
        }
        
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
    }

    // 10. GUARDAR CSV DE PENDIENTES
    if (newUsers.length > 0) {
      const csvContent = 'email,estado\n' + newUsers.map(u => {
        const error = emailErrors.find(e => e.email === u.email);
        return `${u.email},${error ? 'error_email' : 'invitacion_enviada'}`;
      }).join('\n');
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
      isPrivate: true,
      summary: {
        totalUsers: emails.length,
        existingInvited: inviteResults.filter(r => r.status === 'invited').length,
        existingFailed: inviteResults.filter(r => r.status === 'failed').length,
        newUsersEmailed: emailsSent,
        newUsersEmailFailed: emailErrors.length,
        invalidEmails: invalidEmails.length,
        pendingForBot: emailsSent // Los que el bot va a procesar
      },
      existingUsers: inviteResults,
      newUsers: newUsers.map(u => ({
        email: u.email,
        emailSent: !emailErrors.find(e => e.email === u.email),
        addedToBotQueue: !emailErrors.find(e => e.email === u.email)
      })),
      emailErrors: emailErrors,
      invalidEmails: invalidEmails,
      message: `✅ Canal privado creado. ${inviteResults.filter(r => r.status === 'invited').length} usuarios añadidos. ${emailsSent} esperando unirse (bot automático los añadirá).`,
      botInfo: '🤖 Los nuevos usuarios serán añadidos automáticamente al canal cuando se unan al workspace (chequeo cada 30 min)'
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
// ENDPOINT: DESCARGAR PLANTILLA EXCEL SIMPLIFICADA
// ==========================================
app.get('/api/download-template', (req, res) => {
  try {
    const workbook = XLSX.utils.book_new();
    
    // Nueva plantilla simplificada
    const data = [
      ['NOMBRE_CANAL', 'semillero-2025'],
      ['DESCRIPCION', 'Canal para el programa de formación Semillero'],
      ['CORREO', ''],
      ['ejemplo1@correo.com', ''],
      ['ejemplo2@correo.com', ''],
      ['ejemplo3@correo.com', '']
    ];
    
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    worksheet['!cols'] = [{ wch: 25 }, { wch: 50 }];
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Plantilla Slack');
    
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Disposition', 'attachment; filename=plantilla_slack_unicatolica.xlsx');
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
  const config = {
    slackToken: !!process.env.SLACK_BOT_TOKEN,
    workspaceInvite: !!process.env.WORKSPACE_INVITE_LINK,
    sendgrid: !!process.env.SENDGRID_API_KEY,
    emailFrom: !!process.env.EMAIL_FROM
  };
  
  const allConfigured = config.slackToken && config.workspaceInvite && config.sendgrid && config.emailFrom;
  
  res.json({
    status: allConfigured ? 'ok' : 'incomplete',
    configuration: config,
    timestamp: new Date().toISOString()
  });
});

// ==========================================
// ENDPOINT: EJECUTAR BOT MONITOR (para cron)
// ==========================================
app.get('/api/run-monitor', async (req, res) => {
  try {
    console.log("🕒 Cron externo activó el monitor...");
    await checkPendingInvites();
    res.json({ success: true, message: "Monitor ejecutado correctamente" });
  } catch (error) {
    console.error("❌ Error ejecutando monitor:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});



// Crear carpetas necesarias
if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
if (!fs.existsSync('public')) fs.mkdirSync('public');

// Iniciar servidor
app.listen(PORT, () => {
  console.log('\n✨ ========================================');
  console.log(`   🌱 Sistema Automático SLACK - Unicatólica`);
  console.log(`   🚀 http://localhost:${PORT}`);
  console.log('========================================\n');
  console.log('🔧 Verificando configuración del sistema:');
  console.log(`   Token Slack: ${process.env.SLACK_BOT_TOKEN ? '✅' : '❌ NO CONFIGURADO'}`);
  console.log(`   Link Workspace: ${process.env.WORKSPACE_INVITE_LINK ? '✅' : '❌ NO CONFIGURADO'}`);
  console.log(`   SendGrid API: ${process.env.SENDGRID_API_KEY ? '✅' : '❌ NO CONFIGURADO'}`);
  console.log(`   Email From: ${process.env.EMAIL_FROM ? '✅' : '❌ NO CONFIGURADO'}`);
  console.log('\n📋 Formato del Excel simplificado:');
  console.log('   Fila 1: NOMBRE_CANAL | nombre-del-canal');
  console.log('   Fila 2: DESCRIPCION | Descripción del canal');
  console.log('   Fila 3: CORREO');
  console.log('   Fila 4+: Un correo por fila\n');
});

