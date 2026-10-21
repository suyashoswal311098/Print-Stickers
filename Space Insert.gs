function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const ss = e.source;
  
  // 1) Only proceed if it really was a check in Print Console!K2
  if (
    sheet.getName() !== 'Print Console' ||
    range.getA1Notation() !== 'K2' ||
    e.value !== 'TRUE'
  ) return;  
  
  // 2) Untick immediately so that we don’t retrigger while running
  range.setValue(false);
  
  // 2b) Activate B2 before giving feedback
  sheet.getRange('B2').activate();
  
  // 3) Give immediate feedback
  ss.toast('Running insert3…', 'Please wait', 3);
  
  // 4) Prevent accidental double-runs if the user clicks twice
  const lock = LockService.getScriptLock();
  try {
    // wait up to 10 s to acquire
    lock.waitLock(5000);
    
    // 5) Call your heavy function
    insert3(e); 
    
    // 6) Final feedback when done
    ss.toast('✅ insert3 complete!', 'Done', 3);
    
  } catch (err) {
    // if lock fails / insert3 crashes
    ss.toast('⚠️ insert3 failed: ' + err.message, 'Error', 5);
    Logger.log(err);
  } finally {
    lock.releaseLock();
  }
}
