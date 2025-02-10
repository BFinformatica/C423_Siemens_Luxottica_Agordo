Se è master scrive 1 sulla tag del comunicator. Se è slave scrive 0.

Nel file RoleMonitor.ini inserire
[TAG COMUNICATOR PER RUOLO (0 slave, 1 master)]
[RUOLO PREDREFINITO (MASTER o SLAVE) DEL PC LOCALE]
[IP DEL PC MASTER X PING (se locale master mettere 127.0.0.1)]
[PATH FILE WATCHDOG DEL PC MASTER PREDEFINITO] 
	potrebbe gestire più di un file watchdog mettendo i precorsi nelle righe seguenti ma questa funz non è stata testata


Nel file RoleMonitor_PartnerMonitoring.ini inserire
[NOME TAG COMUNICATOR PER DIGITALE STATO PARTNER (Es. 1 DI99)]
[IP PC PARTNER]
[PATH FILE WATCHDOG DEL PC PARTNER]