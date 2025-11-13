require('dotenv').config({ path: 'canal.env' });
const { WebClient } = require('@slack/web-api');
const fs = require('fs');
const path = require('path');

// ==========================================
// CONFIGURACIÓN
// ==========================================
const slack = new WebClient(process.env.SLACK_BOT_TOKEN);
const PENDING_FILE = path.join(__dirname, 'pendientes.json');
const CHECK_INTERVAL = 30 * 60 * 1000; // 30 minutos en milisegundos

console.log('🔧 Configuración Bot de Monitoreo:');
console.log(`   Archivo pendientes: ${PENDING_FILE}`);
console.log(`   Intervalo de chequeo: ${CHECK_INTERVAL / 60000} minutos`);
console.log(`   Token configurado: ${process.env.SLACK_BOT_TOKEN ? '✅' : '❌'}\n`);

// ==========================================
// FUNCIONES PARA MANEJAR ARCHIVO PENDIENTES
// ==========================================

/**
 * Carga la lista de usuarios pendientes desde el archivo JSON
 * @returns {Array} Lista de usuarios pendientes
 */
function loadPending() {
  try {
    if (fs.existsSync(PENDING_FILE)) {
      const data = fs.readFileSync(PENDING_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('❌ Error leyendo pendientes:', error.message);
  }
  return [];
}

/**
 * Guarda la lista de usuarios pendientes en el archivo JSON
 * @param {Array} pending - Lista de usuarios pendientes
 */
function savePending(pending) {
  try {
    fs.writeFileSync(PENDING_FILE, JSON.stringify(pending, null, 2), 'utf8');
  } catch (error) {
    console.error('❌ Error guardando pendientes:', error.message);
  }
}

/**
 * Agrega un usuario a la lista de pendientes
 * @param {string} email - Email del usuario
 * @param {string} channelId - ID del canal de Slack
 * @param {string} channelName - Nombre del canal
 */
function addPending(email, channelId, channelName) {
  const pending = loadPending();
  
  // Evitar duplicados
  const exists = pending.find(p => 
    p.email === email && p.channelId === channelId
  );
  
  if (exists) {
    console.log(`⚠️  Ya existe en pendientes: ${email} → #${channelName}`);
    return;
  }
  
  pending.push({
    email,
    channelId,
    channelName,
    addedAt: new Date().toISOString()
  });
  
  savePending(pending);
  console.log(`📋 Agregado a pendientes: ${email} → #${channelName}`);
}

/**
 * Limpia usuarios pendientes que llevan más de 7 días
 */
function cleanOldPending() {
  const pending = loadPending();
  const oneWeekAgo = Date.now() - (7 * 24 * 60 * 60 * 1000);
  
  const cleaned = pending.filter(p => {
    const addedTime = new Date(p.addedAt).getTime();
    return addedTime > oneWeekAgo;
  });
  
  const removed = pending.length - cleaned.length;
  
  if (removed > 0) {
    savePending(cleaned);
    console.log(`🗑️  Eliminados ${removed} pendientes antiguos (>7 días)`);
  }
}

// ==========================================
// BOT DE MONITOREO - FUNCIÓN PRINCIPAL
// ==========================================

/**
 * Revisa y procesa usuarios pendientes
 */
async function checkPendingInvites() {
  console.log('\n🔍 ========================================');
  console.log('   BOT DE MONITOREO EJECUTÁNDOSE');
  console.log(`   Hora: ${new Date().toLocaleString('es-ES')}`);
  console.log('========================================');
  
  // Limpiar pendientes antiguos
  cleanOldPending();
  
  const pending = loadPending();
  
  if (pending.length === 0) {
    console.log('   ℹ️  No hay usuarios pendientes de procesar');
    console.log('========================================\n');
    return;
  }
  
  console.log(`   📋 Usuarios pendientes: ${pending.length}\n`);
  
  const stillPending = [];
  let processed = 0;
  let invited = 0;
  let alreadyInChannel = 0;
  
  for (const item of pending) {
    console.log(`   🔍 Revisando: ${item.email}`);
    
    try {
      // Buscar usuario por email en Slack
      const user = await slack.users.lookupByEmail({ 
        email: item.email 
      });
      
      // ¡Usuario encontrado! Ya está en el workspace
      console.log(`      ✅ Usuario encontrado en workspace`);
      
      // Intentar invitar al canal privado
      try {
        await slack.conversations.invite({
          channel: item.channelId,
          users: user.user.id
        });
        
        console.log(`      ✅ Invitado exitosamente a #${item.channelName}`);
        invited++;
        processed++;
        
        // NO agregar a stillPending (ya procesado)
        
      } catch (inviteError) {
        if (inviteError.data?.error === 'already_in_channel') {
          console.log(`      ✅ Ya estaba en el canal #${item.channelName}`);
          alreadyInChannel++;
          processed++;
          // No agregar a stillPending (ya está en el canal)
        } else {
          console.log(`      ⚠️  Error invitando: ${inviteError.data?.error}`);
          stillPending.push(item); // Mantener para reintentar
        }
      }
      
    } catch (error) {
      if (error.data?.error === 'users_not_found') {
        // Usuario aún no se ha unido al workspace
        console.log(`      ⏳ Usuario aún no se unió al workspace`);
        stillPending.push(item); // Mantener en lista
      } else {
        console.log(`      ❌ Error: ${error.data?.error || error.message}`);
        stillPending.push(item); // Mantener para revisar después
      }
    }
    
    // Pausa entre requests para evitar rate limits
    await new Promise(resolve => setTimeout(resolve, 300));
  }
  
  // Actualizar archivo de pendientes
  savePending(stillPending);
  
  console.log(`\n   📊 RESUMEN:`);
  console.log(`      ✅ Procesados exitosamente: ${processed}`);
  console.log(`      👥 Nuevos invitados: ${invited}`);
  console.log(`      ✓  Ya estaban en canal: ${alreadyInChannel}`);
  console.log(`      ⏳ Aún pendientes: ${stillPending.length}`);
  console.log('========================================\n');
}

// ==========================================
// INICIAR BOT DE MONITOREO
// ==========================================

/**
 * Inicia el bot de monitoreo en modo continuo
 */
async function startMonitorBot() {
  console.log('\n🤖 ========================================');
  console.log('   BOT DE MONITOREO INICIADO');
  console.log('========================================');
  console.log(`   📅 Fecha: ${new Date().toLocaleString('es-ES')}`);
  console.log(`   ⏱️  Intervalo: cada ${CHECK_INTERVAL / 60000} minutos`);
  console.log(`   📁 Archivo: ${path.basename(PENDING_FILE)}`);
  console.log('========================================\n');
  
  // Ejecutar inmediatamente al iniciar
  console.log('🚀 Ejecutando primer chequeo...\n');
  await checkPendingInvites();
  
  // Luego ejecutar cada intervalo definido
  setInterval(async () => {
    try {
      await checkPendingInvites();
    } catch (error) {
      console.error('❌ Error en bot de monitoreo:', error);
    }
  }, CHECK_INTERVAL);
}

// ==========================================
// EJECUCIÓN
// ==========================================

// Si se ejecuta directamente (node bot-monitor.js)
if (require.main === module) {
  startMonitorBot().catch(error => {
    console.error('❌ Error fatal en bot:', error);
    process.exit(1);
  });
}

// Exportar funciones para usar en server.js
module.exports = {
  addPending,
  loadPending,
  savePending,
  checkPendingInvites,
  startMonitorBot
};