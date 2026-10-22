// Saby's XLSX uses labels, while cached API events use the numeric RPC codes.
// Do not use a generic ordinal or treat an unknown label as an exit.
export function sabyExportDirection(label){
  const value=String(label??'').trim();
  if(value==='Вход')return 1;
  if(value==='Выход')return 0;
  throw new Error(`Неизвестное направление проходной: ${value}`);
}
