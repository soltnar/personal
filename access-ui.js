// Fixed venue directory: independent of report period and selected filters.
const accessVenues = ['Б. Покровская, 59 Самурай','Б. Покровская, 63 Самурай','Белинского, 61 Рибс','Белинского, 61 Самурай','Веденяпина, 1А Самурай','Верхне-Печерская, 14Б Самурай','Волжская, 13 Самурай','Кварц, Гагарина, 224А','Коминтерна 166 CALL CENTRE','Коминтерна, 115 Самурай','Ленина, 36 Самурай','Ленина, 64 Ударник','Моторный, 2/1 доставка Самурай','Октябрьская, 1 Винедо','Октября, 2 Самурай','Ошарская, 8А 19','Парк Швейцария Самурай','Циолковского, 19А Самурай','Фабрика кухни'];
function renderAccessVenues() {
  const select=document.getElementById('accessRestaurant'),selected=new Set([...select.selectedOptions].map(o=>o.value)),list=document.getElementById('accessVenueChecklist');
  select.replaceChildren();list.replaceChildren();
  for(const name of accessVenues){
    const option=document.createElement('option');option.value=name;option.textContent=name;option.selected=selected.has(name);select.append(option);
    const label=document.createElement('label');label.className='restaurantChoice';const check=document.createElement('input');check.type='checkbox';check.checked=option.selected;
    check.onchange=()=>{option.selected=check.checked;updateAccessVenueCount();};
    const text=document.createElement('span');text.textContent=name==='Фабрика кухни'?`${name} · производство`:name;label.append(check,text);list.append(label);
  }
  updateAccessVenueCount();
}
function updateAccessVenueCount(){document.getElementById('accessVenueCount').textContent=`Выбрано: ${document.getElementById('accessRestaurant').selectedOptions.length} из ${accessVenues.length}`;}
document.getElementById('accessAllVenues').onclick=()=>{for(const option of document.getElementById('accessRestaurant').options)option.selected=option.value!=='Фабрика кухни';renderAccessVenues();};
document.getElementById('accessNoVenues').onclick=()=>{for(const option of document.getElementById('accessRestaurant').options)option.selected=false;renderAccessVenues();};
renderAccessVenues();
