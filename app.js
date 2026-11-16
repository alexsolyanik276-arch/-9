/* Квартал Менеджмент — интерактивный пилот v11.5.0 · безопасный офлайн-помощник без ML-модели. */
(() => {
  const $ = (s, root=document) => root.querySelector(s);
  const $$ = (s, root=document) => [...root.querySelectorAll(s)];
  const esc = (v='') => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const norm = v => String(v ?? '').toLowerCase().replace(/ё/g,'е').replace(/[^a-zа-я0-9]+/gi,' ').trim();
  function normalizeChannel(value){
    const raw=String(value??'').trim();const n=norm(raw);
    if(!n)return '';
    if(n.includes('портал'))return 'портал';
    if(n.includes('почт')||n.includes('email')||n.includes('e mail'))return 'почта';
    if((n.includes('общий')||n.includes('общая'))&&n.includes('телефон'))return 'общий телефон стойки регистрации';
    if(n.includes('стойка')&&(/(^| )a$/.test(n)||n.includes('корпус a')))return 'стойка регистрации корпуса A';
    if(n.includes('стойка')&&(/(^| )b$/.test(n)||n.includes('корпус b')))return 'стойка регистрации корпуса B';
    if(n.includes('стойка'))return 'стойка регистрации';
    return raw;
  }
  function channelFilterKey(value){
    const channel=normalizeChannel(value);
    if(!channel)return 'missing';
    if(channel==='портал')return 'portal';
    if(channel==='почта')return 'email';
    if(channel==='стойка регистрации корпуса A')return 'desk-a';
    if(channel==='стойка регистрации корпуса B')return 'desk-b';
    if(channel==='общий телефон стойки регистрации')return 'desk-phone';
    return 'other';
  }
  function appealBuilding(a){
    const fromChannel=channelFilterKey(a.channel);const work=state.works?.find(w=>w.id===a.workId);
    return String(a.building||work?.building||(fromChannel==='desk-a'?'A':fromChannel==='desk-b'?'B':'')).toUpperCase();
  }
  function appealRoom(a){
    const work=state.works?.find(w=>w.id===a.workId);
    return String(a.room||work?.rooms?.[0]||work?.location||'').trim();
  }
  function locationMissing(a){
    return !appealBuilding(a)||!appealRoom(a);
  }
  function locationMissingText(a){
    const missing=[];
    if(!appealBuilding(a))missing.push('Корпус');
    if(!appealRoom(a))missing.push('Помещение');
    return missing;
  }
  function roomInputValue(room,building=''){let value=String(room||'').trim();if(!value)return '';const b=String(building||'').trim().toUpperCase();if(b)value=value.replace(new RegExp('^'+b+'\\s*[-–]?\\s*','i'),'');value=value.replace(/^корпус\s+[AB]\s*/i,'');return value;}
  function canonicalRoomValue(room,building=''){let value=String(room||'').trim().replace(/^корпус\s+[AB]\s*/i,'');const b=String(building||'').trim().toUpperCase();const coded=value.match(/^([AB])\s*[-–]?\s*(\d{3})$/i);if(coded)return (b||coded[1].toUpperCase())+'-'+coded[2];if(b&&/^\d{3}$/.test(value))return b+'-'+value;return value;}
  const state = {data:null, appeals:[], works:[], spaces:[], rules:[], view:'overview', selectedAppeal:null, selectedProblemWork:null, selectedMessage:null, selectedWork:null, query:'', filter:'all', appealChannelFilter:'all', appealBuildingFilter:'all', overviewBuildingFilter:'all', workSort:'newest', workStatusFilter:'all', sourceName:'case_data.xlsx', revision:'', appRevision:'', appState:{tenantQuestions:{},tenantMessages:{},workHistory:{}}, aiStatus:null, aiBusy:false, aiAnswer:{}, aiAnalysis:{}, history:[]};
  let saveTimer=null, saving=false, saveAgain=false, polling=false;

  function firstHeader(rows, predicate){
    for(let i=0;i<Math.min(rows.length,12);i++){
      const cells=(rows[i]||[]).map(x=>norm(x));
      if(predicate(cells)) return i;
    }
    return -1;
  }
  function col(headers, names){
    const h=headers.map(norm);
    for(const n of names){const idx=h.indexOf(norm(n));if(idx>=0)return idx;}
    for(let i=0;i<h.length;i++){if(names.some(n=>h[i].includes(norm(n))))return i;}
    return -1;
  }
  function cell(row, idx){return idx<0?'':String(row?.[idx]??'').trim();}
  function parseRaw(raw){
    const entries=Object.entries(raw||{});
    const sheet=(terms)=>entries.find(([n])=>terms.some(t=>norm(n).includes(norm(t))))?.[1]||[];
    const ar=sheet(['Обращения','Заявки']);
    const sp=sheet(['Помещения','Справочник помещений']);
    const wo=sheet(['Работы','Исполнители']);
    const ru=sheet(['Правила']);
    const ah=firstHeader(ar,h=>h.includes('время')&&h.includes('заявитель')&&h.includes('текст обращения'));
    const sh=firstHeader(sp,h=>h.includes('корпус')&&h.includes('помещение')&&h.includes('зона обслуживания'));
    const wh=firstHeader(wo,h=>h.includes('номер работы')||h.includes('работы')&&h.includes('исполнитель')&&h.includes('статус'));
    const rh=firstHeader(ru,h=>h.includes('тема')&&h.includes('правило'));
    if(ah<0||sh<0||wh<0){throw new Error('В Excel не найдены обязательные листы/колонки: «Обращения», «Помещения», «Работы». Формат листов должен быть близок к шаблону кейса.');}
    const ahd=ar[ah], shd=sp[sh], whd=wo[wh];
    const appealRows=ar.slice(ah+1).filter(r=>r&&r.some(v=>String(v??'').trim()));
    const spaceRows=sp.slice(sh+1).filter(r=>r&&r.some(v=>String(v??'').trim()));
    const workRows=wo.slice(wh+1).filter(r=>r&&r.some(v=>String(v??'').trim()));
    const ruleRows=rh>=0?ru.slice(rh+1).filter(r=>r&&r.some(v=>String(v??'').trim())):[];
    const appeals=appealRows.map((r,i)=>{
      const clarCol=col(ahd,['Требует уточнения']);
      const clarValue=cell(r,clarCol);
      const missingText=cell(r,col(ahd,['Что уточнить']));
      const routeValue=cell(r,col(ahd,['Маршрут']));
      return {
        id:cell(r,col(ahd,['№','Номер','ID']))||`О-${String(i+1).padStart(2,'0')}`,
        time:cell(r,col(ahd,['Время','Дата и время'])),channel:cell(r,col(ahd,['Канал','Источник'])),
        portal:cell(r,col(ahd,['Номер на портале','Номер портала','Portal ID'])) || ((cell(r,col(ahd,['Текст обращения','Текст','Описание'])).match(/\bP\d+\b/i)||[])[0]||''),
        reporter:cell(r,col(ahd,['Заявитель','От кого','Автор'])),text:cell(r,col(ahd,['Текст обращения','Текст','Описание'])),
        building:cell(r,col(ahd,['Корпус'])),room:cell(r,col(ahd,['Помещение'])),zone:cell(r,col(ahd,['Зона обслуживания','Зона'])),
        workId:cell(r,col(ahd,['Связанная работа','№ работы'])),priority:cell(r,col(ahd,['Приоритет'])),
        storedClarification:clarValue===''?null:/^(да|yes|true|1)$/i.test(clarValue),
        storedMissing:missingText?missingText.split(/[;|]/).map(x=>x.trim()).filter(Boolean):null,
        aiSummary:cell(r,col(ahd,['Сводка ИИ'])),responseDraft:cell(r,col(ahd,['Ответ заявителю'])),
        route:routeValue==='tenant'||norm(routeValue).includes('арендатор')?'tenant':'',routeText:routeValue,
        manualUnlinked:/^(да|yes|true|1)$/i.test(cell(r,col(ahd,['Ручная привязка']))),
        linkReason:cell(r,col(ahd,['Причина связи'])),duplicateOf:cell(r,col(ahd,['Дубликат'])),changeLog:cell(r,col(ahd,['История изменений']))
      };
    });
    const spaces=spaceRows.map(r=>({
      building:cell(r,col(shd,['Корпус'])),room:cell(r,col(shd,['Помещение'])),alias:cell(r,col(shd,['Как ещё называют','Как еще называют','Альтернативное название'])),
      floor:cell(r,col(shd,['Этаж'])),zone:cell(r,col(shd,['Зона обслуживания','Зона'])),tenant:cell(r,col(shd,['Арендатор'])),owner:cell(r,col(shd,['Кто обслуживает','Обслуживает']))
    }));
    const works=workRows.map(r=>({
      id:cell(r,col(whd,['№ работы','Номер работы','Работа'])),building:cell(r,col(whd,['Корпус'])),location:cell(r,col(whd,['Помещение или зона','Помещение','Зона'])),zone:cell(r,col(whd,['Зона обслуживания','Зона'])),
      executor:cell(r,col(whd,['Исполнитель'])),status:cell(r,col(whd,['Статус'])),updated:cell(r,col(whd,['Статус обновлён','Статус обновлен','Обновлено'])),
      description:cell(r,col(whd,['Описание'])),related:cell(r,col(whd,['Связанные обращения','Обращения'])),
      priority:cell(r,col(whd,['Приоритет'])),previousWork:cell(r,col(whd,['Предыдущая работа'])),source:cell(r,col(whd,['Источник'])),aiSummary:cell(r,col(whd,['Сводка ИИ']))
    }));
    const rules=ruleRows.map(r=>({id:cell(r,col(ru[rh],['№','Номер'])),topic:cell(r,col(ru[rh],['Тема'])),text:cell(r,col(ru[rh],['Правило']))}));
    if(!appeals.length) throw new Error('Лист «Обращения» не содержит строк с обращениями.');
    return {appeals,spaces,works,rules};
  }

  function getSpaceForText(text, reporter, spaces){
    const n=norm(text), rep=norm(reporter);
    const id=(n.match(/\b([ab])\s*[-–]?\s*(\d{3})\b/i)||[]);
    if(id.length){const building=id[1].toUpperCase(), room=building+'-'+id[2];const found=spaces.find(s=>s.building.toUpperCase()===building&&s.room.toUpperCase()===room);return found||{building,room,zone:'',owner:'управляющая компания'};}
    if(n.includes('у кофейн')||n.includes('кофейни')) return spaces.find(s=>s.building.toUpperCase()==='B'&&norm(s.room).includes('холл'))||null;
    if(n.includes('холл')&&(n.includes('корпус а')||n.includes('холл а')||n.includes('у лифт'))) return spaces.find(s=>s.building.toUpperCase()==='A'&&norm(s.room).includes('холл'))||null;
    if(n.includes('холл')&&n.includes('корпус b')) return spaces.find(s=>s.building.toUpperCase()==='B'&&norm(s.room).includes('холл'))||null;
    // По разговорному названию из справочника; арендатора используем только как дополнительный контекст.
    const aliases=spaces.filter(s=>s.alias).sort((a,b)=>norm(b.alias).length-norm(a.alias).length);
    for(const s of aliases){
      const alias=norm(s.alias), chunks=alias.split(' ').filter(w=>w.length>3);
      if(alias&&n.includes(alias)) return s;
      if(chunks.length>=2&&chunks.every(w=>n.includes(w))) return s;
    }
    if(n.includes('серверн')&&rep.includes('дельта софт')) return spaces.find(s=>s.building==='A'&&s.room==='A-206')||null;
    if(n.includes('переговорн')&&n.includes('ока')) return spaces.find(s=>s.building==='A'&&s.room==='A-310')||null;
    if(n.includes('у лифт')&&n.includes('корпус a')) return spaces.find(s=>s.building==='A'&&norm(s.room).includes('холл'))||null;
    return null;
  }
  function categoryOf(text){
    const n=norm(text);
    if(/створк|окно.*не закры|не закры.*окн|окн.*не откры|не откры.*окн/.test(n))return 'окно';
    if(/ручк|двер/.test(n))return 'дверь';
    if(/светильник|не горит свет|ламп/.test(n))return 'освещение';
    if(/розет/.test(n))return 'электрика';
    if(/серверн|кондиционер|гудит/.test(n))return 'кондиционер';
    if(/холод|прохлад|температур|градус|батаре|дует/.test(n))return 'температура';
    if(/раковин|капает|теч|сантех/.test(n))return 'сантехника';
    if(/поскольз|разлит|пролит|вод.*пол|пятн.*пол|полу/.test(n))return 'пролив на полу';
    if(/пятн|уборк|гряз/.test(n))return 'уборка';
    return 'прочее';
  }
  function shortIssue(cat){return ({'температура':'холодно в помещении','дверь':'неисправность двери','освещение':'не работает светильник','электрика':'не работают розетки','окно':'не закрывается окно','кондиционер':'шум кондиционера','сантехника':'течь / сантехника','пролив на полу':'вода на полу','уборка':'нужна уборка','прочее':'требуется осмотр'})[cat]||cat;}
  function statusClass(status){const s=norm(status);if(s.includes('в работе'))return 'status-progress';if(s.includes('принято'))return 'status-accepted';if(s.includes('открыто'))return 'status-open';if(s.includes('закрыто'))return 'status-closed';if(s.includes('у арендат'))return 'status-tenant';return 'status-clarify';}
  function statusLabel(s){return s||'не задан';}
  function isResolvedStatus(status){const n=norm(status);return ['закрыто','решено','решена','решен','выполнено','выполнена','завершено','завершена','устранено','устранена'].some(x=>n===x||n.startsWith(x+' '));}
  function isResolvedGroup(group){
    if(!group.work||!isResolvedStatus(group.status))return false;
    // Повторное сообщение по закрытой работе остаётся в активном обзоре,
    // пока диспетчер не уточнил, когда неисправность возникла снова.
    const recurrenceNeedsClarification=(group.appeals||[]).some(a=>a.needsClarification&&(a.missing||[]).some(x=>/повтор|снова/i.test(x)));
    return !recurrenceNeedsClarification;
  }
  function isResolvedThanks(text){return /устранили|починили|выполнено|убрали|спасибо|решено/.test(norm(text));}
  function isRecurrence(text){return /снова|опять|по прежнему|по прежнему|все еще|всё ещё|до сих пор|так и не|не работает|не устранили/.test(norm(text));}
  function roomList(location){const m=[...String(location||'').matchAll(/\b([AB])\s*[-–]?\s*(\d{3})\b/gi)];return [...new Set(m.map(x=>x[1].toUpperCase()+'-'+x[2]))];}
  function normalizeWork(w,spaces){
    const ids=roomList(w.location);
    let space=ids.length?spaces.find(s=>norm(s.room)===norm(ids[0])):null;
    if(!space&&w.location){
      const exact=spaces.find(s=>norm(s.room)===norm(w.location)&&(!w.building||s.building.toUpperCase()===w.building.toUpperCase()));
      space=exact||null;
    }
    if(!space&&norm(w.location).includes('холл')) space=spaces.find(s=>s.building.toUpperCase()===(w.building||'').toUpperCase()&&norm(s.room).includes('холл'));
    const category=categoryOf(w.description+' '+w.location);
    const building=w.building||space?.building||'';
    const zone=w.zone||space?.zone||'';
    const location=w.location||(ids.join(', '))||(space?.room||'');
    return {...w,building,location,zone,rooms:ids.length?ids:(space?[space.room]:[]),space:space||null,category,appealIds:(w.related||'').split(/[,;]+/).map(x=>x.trim()).filter(Boolean),source:w.source||'existing',needsClarification:false,notes:[]};
  }
  function buildModel(parsed){
    const spaces=parsed.spaces;
    const works=parsed.works.map(w=>normalizeWork(w,spaces));
    const appeals=parsed.appeals.map(a=>{
      let sp=a.room?spaces.find(s=>norm(s.room)===norm(a.room)&&(!a.building||s.building.toUpperCase()===a.building.toUpperCase())):null;
      if(!sp)sp=getSpaceForText(a.text,a.reporter,spaces);
      const cat=categoryOf(a.text);
      const storedMissing=(a.storedMissing||[]).filter(m=>{
        const n=norm(m);
        if(n.includes('корпус'))return !a.building&&!sp?.building;
        if(n.includes('помещен'))return !a.room&&!sp?.room;
        if(n.includes('заявител')||n.includes('контакт'))return !a.reporter||norm(a.reporter)==='не записан';
        return true;
      });
      return {...a,building:a.building||sp?.building||'',room:a.room||sp?.room||'',zone:a.zone||sp?.zone||'',owner:sp?.owner||'',category:cat,issue:shortIssue(cat),workId:a.workId||'',needsClarification:!!a.storedClarification&&storedMissing.length>0,missing:storedMissing,route:a.route||'',priority:a.priority||(/поскольз|гости жалуются|до 13 00/.test(norm(a.text))?'Высокий':'Обычный'),source:'incoming',duplicateOf:a.duplicateOf||'',aiSummary:a.aiSummary||'',responseDraft:a.responseDraft||''};
    });
    const appById=new Map(appeals.map(a=>[a.id,a]));
    const workById=new Map(works.map(w=>[w.id,w]));
    const attach=(a,w,reason='')=>{a.workId=w.id;if(!w.appealIds.includes(a.id))w.appealIds.push(a.id);if(reason)a.linkReason=reason;};
    // Прежде всего сохраняем связи из выгрузки исполнителей и из явного поля обращения.
    // Если в строке обращения явно задана другая работа, она имеет приоритет над старой связью в листе работ.
    works.forEach(w=>{w.appealIds=(w.appealIds||[]).filter(id=>{const a=appById.get(id);return !a||!a.workId||a.workId===w.id;});});
    works.forEach(w=>w.appealIds.forEach(id=>{const a=appById.get(id);if(a)attach(a,w,'Связь из исходной системы работ');}));
    appeals.forEach(a=>{if(a.workId){const w=workById.get(a.workId);if(w)attach(a,w,a.linkReason||'Связь из карточки обращения');}});
    let next=works.reduce((n,w)=>Math.max(n,Number((w.id.match(/\d+/)||[])[0]||0),n),0)+1;
    const makeWork=(a,old=null)=>{
      const id='РП-'+String(next++).padStart(2,'0');
      const w={id,building:a.building,location:a.room||'Помещение уточняется',zone:a.zone||'—',executor:'Не назначен',status:'открыто',updated:workStamp(),description:a.issue,rooms:a.room?[a.room]:[],space:spaces.find(s=>s.room===a.room)||null,category:a.category,appealIds:[],source:'pilot',needsClarification:false,notes:[],previousWork:old?.id||''};
      works.push(w);workById.set(id,w);attach(a,w,old?`Повтор после ${old.id}`:'Новая неисправность');return w;
    };
    for(const a of appeals){
      if(a.workId||a.manualUnlinked) continue;
      const n=norm(a.text);
      // Дубли одного портального номера сохраняются как отдельные входящие, но не создают работу.
      const dup=appeals.find(x=>x!==a&&a.portal&&x.portal===a.portal&&x.workId);
      if(dup){a.duplicateOf=dup.id;const w=workById.get(dup.workId);if(w)attach(a,w,'Повтор того же номера портала');continue;}
      if(a.owner&&norm(a.owner).includes('арендатор')){a.route='tenant';a.needsClarification=false;a.routeText=`Обслуживает арендатор: ${a.owner}`;continue;}
      const matches=works.filter(w=>{
        const active=norm(w.status)!=='закрыто';
        if(!active) return false;
        if(a.category==='температура') return a.zone&&w.zone===a.zone&&w.category==='температура';
        if(!a.room) return false;
        const sameRoom=(w.rooms||[]).includes(a.room)||w.location===a.room;
        return sameRoom&&w.category===a.category;
      });
      if(matches.length){
        const w=matches[0];attach(a,w,'Та же зона и неисправность');
        if(/встреча|гости/.test(n)){a.priority='Высокий';a.needsClarification=false;}
        continue;
      }
      const closedMatches=works.filter(w=>{
        if(norm(w.status)!=='закрыто') return false;
        if(a.category==='температура') return a.zone&&w.zone===a.zone&&w.category==='температура';
        return a.room&&((w.rooms||[]).includes(a.room)||w.location===a.room)&&w.category===a.category;
      });
      if(closedMatches.length&&isResolvedThanks(a.text)){
        attach(a,closedMatches[0],'Подтверждение результата закрытой работы');continue;
      }
      if(closedMatches.length&&isRecurrence(a.text)){
        a.workId=closedMatches[0].id;a.previousWork=closedMatches[0].id;a.needsClarification=true;
        a.missing.push('Когда проблема появилась повторно');a.linkReason='После закрытия: нужно уточнить время повторения';
        continue;
      }
      if(!a.building||!a.room){a.needsClarification=true;if(!a.building)a.missing.push('Корпус');if(!a.room)a.missing.push('Помещение');continue;}
      const nw=makeWork(a);
      if(a.priority==='Высокий')nw.priority='Высокий';
    }
    // Неизвестного заявителя отмечаем отдельно: это не мешает срочно завести работу, если место ясно.
    for(const a of appeals){
      if(!a.reporter||norm(a.reporter)==='не записан'){a.missing.push('Заявитель / контакт');a.needsClarification=true;}
      if(!a.building&&!a.missing.includes('Корпус'))a.missing.push('Корпус');
      if(!a.room&&!a.missing.includes('Помещение'))a.missing.push('Помещение');
      if(a.workId){const w=workById.get(a.workId);if(w&&norm(w.status)==='закрыто'&&isRecurrence(a.text)&&!isResolvedThanks(a.text)){a.needsClarification=true;if(!a.missing.includes('Когда проблема появилась повторно'))a.missing.push('Когда проблема появилась повторно');}}
    }
    // O-19: работа закрыта, но момент повторения неизвестен — оставить на уточнении, не открывать автоматически.
    // Для заявителя, перенаправленного арендатору, показываем владельца оборудования.
    appeals.forEach(a=>{if(a.route==='tenant')a.routeText='Обслуживает арендатор по справочнику помещений';});
    return {appeals,works,spaces,rules:parsed.rules};
  }

  function loadData(raw,name='case_data.xlsx',revision='',appState=null,appRevision=''){
    const parsed=parseRaw(raw);const model=buildModel(parsed);
    state.data=raw;state.appeals=model.appeals;state.works=model.works;state.spaces=model.spaces;state.rules=model.rules;state.sourceName=name;state.revision=revision||state.revision;state.appRevision=appRevision||state.appRevision;if(appState&&typeof appState==='object')state.appState={...appState,tenantQuestions:appState.tenantQuestions||{},tenantMessages:appState.tenantMessages||{},workHistory:appState.workHistory||{}};
    state.appeals.forEach(a=>{const meta=state.appState.tenantMessages?.[a.id];if(meta){if(meta.responseDraft)a.responseDraft=meta.responseDraft;if(meta.aiSummary)a.aiSummary=meta.aiSummary;}});
    state.history=model.appeals.flatMap(a=>(a.changeLog||'').split('\n').filter(Boolean).map(line=>{const cut=line.indexOf(' · ');return {appealId:a.id,time:cut>=0?line.slice(0,cut):'',text:cut>=0?line.slice(cut+3):line};}));const sourceLabel=document.querySelector('#sourceLabel');if(sourceLabel)sourceLabel.textContent='Книга: data/active.xlsx · '+name;
    // При внешнем обновлении книги сохраняем текущую вкладку, фильтр и поисковый запрос.
    if(state.selectedAppeal&&!state.appeals.some(a=>a.id===state.selectedAppeal))state.selectedAppeal=null;
    if(state.selectedProblemWork&&!state.works.some(w=>w.id===state.selectedProblemWork))state.selectedProblemWork=null;
    if(state.selectedMessage&&!state.appeals.some(a=>a.id===state.selectedMessage))state.selectedMessage=null;
    if(state.selectedWork&&!state.works.some(w=>w.id===state.selectedWork))state.selectedWork=null;
    updateNav();render();
  }
  function updateNav(){
    const count=state.appeals.filter(locationMissing).length;
    const clarify=$('#navClarifyCount');
    if(clarify){clarify.textContent=count;clarify.classList.toggle('is-empty',count===0);clarify.title=count?`${count} обращений без корпуса или помещения`:'Нет обращений для уточнения';}
  }
  function setSyncStatus(text,kind=''){
    const el=$('#syncStatus');if(!el)return;el.textContent=text;el.className='sync-state '+kind;
  }
  function queueSave(){
    if(saveTimer)clearTimeout(saveTimer);
    setSyncStatus('Сохранение в Excel…','sync-pending');
    saveTimer=setTimeout(saveWorkbook,220);
  }
  async function saveWorkbook(){
    saveTimer=null;
    if(saving){saveAgain=true;return;}
    saving=true;
    const expectedRevision=state.revision;
    const payload={revision:expectedRevision,appRevision:state.appRevision,appeals:state.appeals,works:state.works,appState:state.appState};
    try{
      const response=await fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
      const data=await response.json();
      if(response.status===409){setSyncStatus('Есть внешние изменения','sync-warn');await refreshWorkbook(true);toast('Книга Excel изменилась в другом окне. Обновили данные — повтори последнее действие.','warn');return;}
      if(!response.ok)throw new Error(data.error||'Не удалось сохранить книгу Excel');
      state.revision=data.revision;state.appRevision=data.appRevision||state.appRevision;setSyncStatus('Excel синхронизирован','sync-ok');
    }catch(e){setSyncStatus('Ошибка синхронизации','sync-error');toast(e.message,'error');}
    finally{saving=false;if(saveAgain){saveAgain=false;queueSave();}}
  }
  async function refreshWorkbook(showToast=false){
    try{
      const response=await fetch('/api/bootstrap',{cache:'no-store'});const data=await response.json();
      if(!response.ok)throw new Error(data.error||'Не удалось перечитать Excel');
      if(data.revision!==state.revision||(data.appRevision||'')!==state.appRevision){
        loadData(data.sheets,data.name,data.revision,data.appState,data.appRevision);
        setSyncStatus('Обновлено из Excel','sync-ok');
        if(showToast)toast('Данные обновлены из рабочей книги Excel.');
      }
      return true;
    }catch(e){setSyncStatus('Нет связи с сервером','sync-error');if(showToast)toast(e.message,'error');return false;}
  }
  async function pollWorkbook(){
    if(polling||saving||saveTimer)return;polling=true;
    try{
      const response=await fetch('/api/revision',{cache:'no-store'});if(!response.ok)return;
      const info=await response.json();if((info.revision&&info.revision!==state.revision)||(info.appRevision||'')!==state.appRevision)await refreshWorkbook(true);
    }catch(_){setSyncStatus('Нет связи с сервером','sync-error');}
    finally{polling=false;}
  }
  function toast(msg,type=''){const el=document.createElement('div');el.className='toast '+type;el.textContent=msg;$('#toastRoot').appendChild(el);setTimeout(()=>el.remove(),3400);}
  function badgeStatus(a){
    let base='';
    if(a.route==='tenant') base='<span class="status status-tenant">У арендатора</span>';
    else if(a.workId){const w=state.works.find(x=>x.id===a.workId);base=w?`<span class="status ${statusClass(w.status)}">${esc(statusLabel(w.status))}</span>`:'<span class="status status-clarify">Без работы</span>';}
    else base='';
    const clarify=a.needsClarification?'<span class="status status-clarify">Требует уточнения</span>':'';
    return base+clarify;
  }
  function workTitle(w){const rooms=(w.rooms||[]).join(', ')||w.location||'Помещение не указано';return `${rooms} · ${w.description}`;}
  function issueFor(a){return a.aiSummary||a.issue||shortIssue(a.category);}
  function getAppealWork(a){return state.works.find(w=>w.id===a.workId)||null;}
  function calcGroups(){return state.works.filter(w=>w.appealIds.length>0).length;}
  function workCard(w){const linked=state.appeals.filter(a=>w.appealIds.includes(a.id));const clar=linked.filter(a=>a.needsClarification).length;const loc=(w.rooms||[]).map(r=>roomInputValue(r,w.building)).join(', ')||roomInputValue(w.location,w.building);return `<div class="work-row" data-work="${esc(w.id)}"><div class="work-code">${esc(w.id)}</div><div class="work-info"><div class="work-name">${esc(loc)} · ${esc(w.description)}</div><div class="work-meta"><span>${esc(w.building?'Корпус '+w.building:'Корпус не указан')}</span><span>·</span><span>Зона обслуживания ${esc(w.zone||'не указана')}</span><span>·</span><span>${linked.length} сообщ.</span><span>·</span><span>${esc(w.executor||'Не назначен')}</span>${w.source==='pilot'?'<span>·</span><span class="success-text">Новая работа</span>':''}</div></div><div class="work-end"><span class="status ${statusClass(w.status)}">${esc(w.status)}</span>${clar?`<span class="count-meta">${clar} уточнить</span>`:''}</div></div>`;}
  function attentionItems(){
    return sortedAppeals(state.appeals.filter(locationMissing)).map(a=>{
      const missing=locationMissingText(a);
      const place=[appealBuilding(a)?'Корпус '+appealBuilding(a):'',roomInputValue(appealRoom(a),appealBuilding(a))].filter(Boolean).join(' · ');
      const grammar=missing.length>1?'не указаны':missing[0]==='Корпус'?'не указан':'не указано';
      return {id:a.id,title:`${a.id} · ${missing.join(' и ').toLowerCase()} ${grammar}`,text:`${a.issue||shortIssue(a.category)}${place?' · '+place:''}`};
    });
  }
  function extractTenantQuestion(a){
    const text=String(a?.text||'').trim();if(!text)return '';
    const explicit=text.match(/[^.!?\n]*\?+/g);
    if(explicit?.length)return explicit.map(x=>x.trim()).filter(Boolean).join(' ');
    const sentences=text.split(/(?<=[.!])\s+/);
    const request=sentences.find(x=>/(когда\s+(придут|при[йд]ет|ждать|ожидать|будет)|кто\s+(придет|придёт|ответствен)|почему\s+(не|так)|можно\s+ли|подскажите|просим\s+(сообщить|уточнить)|сообщите\s*,?\s*пожалуйста|жд[её]м\s+новост|хотим\s+узнать|интересует\s+статус)/i.test(x));
    if(request){
      if(/жд[её]м\s+новост/i.test(request))return 'Просит сообщить новости о ходе общей проверки.';
      return request.trim();
    }
    return '';
  }
  function questionInfo(a){
    const stored=state.appState?.tenantQuestions?.[a.id]||{};
    const text=String(stored.text||extractTenantQuestion(a)).trim();
    if(!text)return null;
    return {text,answered:stored.answered===true};
  }
  function groupQuestionInfo(group){
    const questions=group.appeals.map(a=>({appeal:a,question:questionInfo(a)})).filter(x=>x.question);
    const unanswered=questions.filter(x=>!x.question.answered);
    return {questions,unanswered,hasUnanswered:unanswered.length>0};
  }
  function markTenantQuestionAnswered(id,answered=true,feedback=null){
    const a=state.appeals.find(x=>x.id===id);if(!a)return;
    const info=questionInfo(a);if(!info)return;
    state.appState.tenantQuestions=state.appState.tenantQuestions||{};
    state.appState.tenantQuestions[id]={text:info.text,answered:!!answered,updatedAt:new Date().toISOString()};
    addHistory(id,answered?'Вопрос арендатора отмечен как отвеченный.':'Вопрос арендатора снова отмечен как ожидающий ответа.');
    queueSave();
    render();
    toast(feedback||(answered?'Вопрос отмечен как отвеченный':'Вопрос снова ожидает ответа'));
  }
  function appealCountWord(n){const m=n%100; if(m>=11&&m<=14)return 'обращений'; const d=n%10; return d===1?'обращение':(d>=2&&d<=4?'обращения':'обращений');}
  function problemGroups(){
    const groups=new Map();
    state.appeals.forEach(a=>{
      const ownWork=getAppealWork(a);
      const parent=a.duplicateOf?state.appeals.find(x=>x.id===a.duplicateOf):null;
      const parentWork=parent?getAppealWork(parent):null;
      const work=ownWork||parentWork;
      const building=(a.building||work?.building||parent?.building||'').toUpperCase();
      const room=norm(roomInputValue(a.room||work?.location||parent?.room||'',building));
      const zone=norm(a.zone||work?.zone||parent?.zone||'');
      const category=norm(a.category||categoryOf(a.text||''));
      const placeKey=category==='температура'&&zone?('zone-'+zone):(room?('room-'+room):(zone?('zone-'+zone):''));
      let key;
      if(work) key='work:'+work.id;
      else if(parent) key='appeal:'+(parent.duplicateOf||parent.id);
      else if(placeKey) key=['place',building,placeKey,category,a.route==='tenant'?'tenant':'queue'].join(':');
      else key='appeal:'+a.id;
      if(!groups.has(key))groups.set(key,{key,work,appeals:[]});
      const bucket=groups.get(key);if(!bucket.work&&work)bucket.work=work;
      bucket.appeals.push(a);
    });
    const grouped=[...groups.values()].map(g=>{
      const items=sortedAppeals(g.appeals);
      const latest=items[0];
      const work=g.work||getAppealWork(latest);
      const building=work?.building||latest.building||'';
      const room=(work?.rooms||[]).filter(Boolean).map(r=>roomInputValue(r,building)).join(', ')||roomInputValue(work?.location||latest.room||'',building);
      const zone=work?.zone||latest.zone||state.spaces.find(sp=>sp.building===building&&sp.room===latest.room)?.zone||'';
      const issue=latest.issue||shortIssue(latest.category);
      const status=work?.status||(latest.route==='tenant'?'у арендатора':'без работы');
      const unassigned=!work&&latest.route!=='tenant';
      const questions=groupQuestionInfo({...g,appeals:items});
      return { ...g, appeals:items, latest, work, building, room, zone, issue, status, unassigned, updated:appealDateValue(latest.time), category:work?.category||latest.category, questions, hasUnansweredQuestion:questions.hasUnanswered };
    });
    // Самостоятельные проблемы, созданные диспетчером, хранятся как строки листа «Работы» без искусственного обращения арендатора.
    const representedWorkIds=new Set(grouped.map(g=>g.work?.id).filter(Boolean));
    const standalone=state.works.filter(w=>!representedWorkIds.has(w.id)&&!state.appeals.some(a=>a.workId===w.id))
      .map(work=>{
        const building=work.building||'';
        const rooms=(work.rooms||[]).filter(Boolean).map(r=>roomInputValue(r,building));
        const room=rooms.join(', ')||roomInputValue(work.location||'',building);
        const space=state.spaces.find(sp=>building&&sp.building.toUpperCase()===String(building).toUpperCase()&&roomInputValue(sp.room,building)===roomInputValue(room,building));
        const zone=work.zone||space?.zone||'';
        const updated=appealDateValue(work.updated);
        return {key:'work:'+work.id,work,appeals:[],latest:null,building,room,zone,issue:work.description||shortIssue(work.category||'прочее'),status:work.status||'открыто',unassigned:false,updated,category:work.category||categoryOf(work.description||''),questions:{questions:[],unanswered:[],hasUnanswered:false},hasUnansweredQuestion:false,standalone:true};
      });
    return [...grouped,...standalone].sort((a,b)=>Number(b.hasUnansweredQuestion)-Number(a.hasUnansweredQuestion)||b.updated-a.updated);
  }
  function groupMatchesFilter(group,filter){
    if(filter==='status-accepted')return !!group.work&&norm(group.status)==='принято';
    if(filter==='status-open')return !!group.work&&norm(group.status)==='открыто';
    if(filter==='status-progress')return !!group.work&&norm(group.status)==='в работе';
    return true;
  }
  function problemCard(group){
    const linked=group.appeals.length;
    const questionMark=group.hasUnansweredQuestion?`<span class="question-alert question-alert-feed" title="Есть вопрос арендатора без ответа">!</span>`:'';
    // В лентах обзора и решённых проблем показываем только корпус и помещение.
    const place=[group.building?'Корпус '+group.building:'',group.room?'Помещение '+roomInputValue(group.room,group.building):''].filter(Boolean).join(' · ')||'Место уточняется';
    const target=group.latest?.id?`data-appeal="${esc(group.latest.id)}"`:`data-problem-work="${esc(group.work?.id||'')}"`;
    const countText=linked?`${linked} ${appealCountWord(linked)}`:'Создана диспетчером';
    const timeText=group.latest?.time||group.work?.updated||'';
    return `<button type="button" class="problem-card" ${target}>
      <span class="problem-main"><span class="problem-title">${questionMark}${esc(group.issue)}</span><span class="problem-meta"><span>${esc(place)}</span><span>${esc(countText)}</span>${timeText?`<span>Обновлено: ${esc(timeText)}</span>`:''}</span></span>
    </button>`;
  }
  function statusAppealCount(status){return problemGroups().filter(g=>g.work&&norm(g.status)===norm(status)).length;}
  function overviewView(){
    const groups=problemGroups().filter(g=>!isResolvedGroup(g));
    const defs=[
      {label:'Открыто',filter:'status-open',cls:'kpi-open',match:g=>g.work&&norm(g.status)==='открыто'},
      {label:'Принято',filter:'status-accepted',cls:'kpi-accepted',match:g=>g.work&&norm(g.status)==='принято'},
      {label:'В работе',filter:'status-progress',cls:'kpi-progress',match:g=>g.work&&norm(g.status)==='в работе'}
    ];
    const cards=defs.map(d=>{const matched=groups.filter(d.match);return {...d,value:matched.length,hasQuestions:matched.some(g=>g.hasUnansweredQuestion)};});
    const filterLabels={'all':'Все проблемы','status-open':'Открыто','status-accepted':'Принято','status-progress':'В работе'};
    const visibleGroups=groups.filter(g=>groupMatchesFilter(g,state.filter)).sort((a,b)=>Number(b.hasUnansweredQuestion)-Number(a.hasUnansweredQuestion)||b.updated-a.updated);
    return `<div class="page-heading"><div><h1 class="page-title">Обзор</h1></div><div class="heading-actions"><button class="btn btn-primary" id="quickNewBtn">＋ Проблема</button></div></div>
      <div class="kpi-grid status-kpi-grid">${cards.map(c=>`<button type="button" class="kpi-card ${c.cls} ${state.filter===c.filter?'is-active':''}" data-kpi-filter="${c.filter}" aria-pressed="${state.filter===c.filter}" aria-label="Показать проблемы: ${c.label}"><span class="kpi-top"><span>${c.label}</span>${c.hasQuestions?'<span class="kpi-question-mark" title="Есть вопросы арендаторов без ответа">!</span>':''}</span><span class="kpi-value">${c.value}</span></button>`).join('')}</div>
      <section class="panel overview-panel overview-problems-panel"><div class="panel-head"><div><h2 class="panel-title">Проблемы</h2><div class="panel-hint">${visibleGroups.length} · ${esc(filterLabels[state.filter]||'Все проблемы')}${visibleGroups.some(g=>g.hasUnansweredQuestion)?' · сначала вопросы без ответа':''}</div></div></div>
        <div class="overview-scroll-list problem-list">${visibleGroups.length?visibleGroups.map(problemCard).join(''):'<div class="empty-state">Нет проблем по выбранному фильтру</div>'}</div>
      </section>`;
  }
  function resolvedProblemsView(){
    const groups=problemGroups().filter(isResolvedGroup).sort((a,b)=>b.updated-a.updated);
    return `<div class="page-heading"><div><h1 class="page-title">Решённые проблемы</h1><p class="page-subtitle">Проблемы с закрытыми работами</p></div></div>
      <section class="panel overview-panel overview-problems-panel"><div class="panel-head"><div><h2 class="panel-title">Решённые проблемы</h2><div class="panel-hint">${groups.length} ${appealCountWord(groups.length)}</div></div></div>
        <div class="overview-scroll-list problem-list">${groups.length?groups.map(problemCard).join(''):'<div class="empty-state"><strong>Пока нет решённых проблем</strong>Закрытые работы появятся здесь.</div>'}</div>
      </section>`;
  }

  function appealDateValue(value){
    const text=String(value||'').trim();
    if(!text)return 0;
    if(/сейчас|только что|только\s+что/i.test(text))return Number.MAX_SAFE_INTEGER;
    const m=text.match(/(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?\s+(\d{1,2}):(\d{2})/);
    if(m){let year=m[3]?Number(m[3]):2026;if(year<100)year+=2000;return new Date(year,Number(m[2])-1,Number(m[1]),Number(m[4]),Number(m[5])).getTime()||0;}
    const parsed=Date.parse(text);return Number.isFinite(parsed)?parsed:0;
  }
  function sortedAppeals(items){
    // Внутри карточки проблемы связанные сообщения упорядочиваются по времени; вкладка «Входящие» сохраняет порядок Excel.
    return [...items].sort((a,b)=>appealDateValue(b.time)-appealDateValue(a.time));
  }
  function sortedWorks(items){
    const direction=state.workSort==='oldest'?1:-1;
    return [...items].sort((a,b)=>direction*(appealDateValue(a.updated)-appealDateValue(b.updated)));
  }
  function appealRow(a,showStatus=false){
    const building=appealBuilding(a);const location=[building?'Корпус '+building:'',roomInputValue(appealRoom(a),building)].filter(Boolean).join(' · ')||'Место уточняется';
    const issue=(a.issue||shortIssue(a.category)||'Обращение').replace(/^./,x=>x.toUpperCase());
    const q=questionInfo(a);const questionMark=q&&!q.answered?'<span class="question-alert question-alert-feed" title="Вопрос арендатора ожидает ответа">!</span>':'';
    const clarifyLine=showStatus?`<span class="problem-meta problem-clarify-meta"><span>Не указано: ${esc(locationMissingText(a).join(', ').toLowerCase())}</span></span>`:'';
    return `<button type="button" class="problem-card appeal-feed-card ${showStatus?'clarify-appeal-item':''}" data-appeal="${esc(a.id)}"><span class="problem-main"><span class="problem-title">${questionMark}${esc(location)} · ${esc(issue)}</span><span class="problem-meta"><span>${esc(a.reporter||'Заявитель не указан')}</span><span>${esc(a.time||'Время не указано')}</span><span>${esc(a.channel||'Канал не указан')}</span></span>${clarifyLine}</span></button>`;
  }
  function appealsView(onlyClarify=false){
    const filtered=state.appeals.filter(a=>{
      if(onlyClarify&&!locationMissing(a))return false;
      const q=norm(state.query);if(q&&!norm([a.id,a.reporter,a.text,a.building,a.room,a.workId,a.portal,a.channel].join(' ')).includes(q))return false;
      if(!onlyClarify&&state.appealChannelFilter!=='all'&&channelFilterKey(a.channel)!==state.appealChannelFilter)return false;
      if(!onlyClarify){
        const building=appealBuilding(a);
        if(state.appealBuildingFilter==='A'&&building!=='A')return false;
        if(state.appealBuildingFilter==='B'&&building!=='B')return false;
        if(state.appealBuildingFilter==='missing'&&building)return false;
      }
      return true;
    });
    const list=filtered; // Сохраняем исходный порядок строк Excel во входящих.
    const title=onlyClarify?'Нужно уточнить':'Входящие';
    return `<div class="page-heading"><div><h1 class="page-title">${title}</h1></div><div class="heading-actions"><button class="btn btn-primary" id="quickNewBtn">＋ Проблема</button></div></div>
      <div class="table-toolbar"><div class="searchbox"><span>⌕</span><input id="appealSearch" placeholder="Поиск по месту или неисправности" value="${esc(state.query)}"></div>${onlyClarify?'':`<select id="appealChannelFilter" class="filter-select" aria-label="Фильтр по каналу"><option value="all" ${state.appealChannelFilter==='all'?'selected':''}>Все каналы</option><option value="portal" ${state.appealChannelFilter==='portal'?'selected':''}>Портал</option><option value="email" ${state.appealChannelFilter==='email'?'selected':''}>Почта</option><option value="desk-a" ${state.appealChannelFilter==='desk-a'?'selected':''}>Стойка регистрации · корпус A</option><option value="desk-b" ${state.appealChannelFilter==='desk-b'?'selected':''}>Стойка регистрации · корпус B</option><option value="desk-phone" ${state.appealChannelFilter==='desk-phone'?'selected':''}>Общий телефон стойки регистрации</option><option value="missing" ${state.appealChannelFilter==='missing'?'selected':''}>Канал не указан</option></select><select id="appealBuildingFilter" class="filter-select" aria-label="Фильтр по корпусу"><option value="all" ${state.appealBuildingFilter==='all'?'selected':''}>Все корпуса</option><option value="A" ${state.appealBuildingFilter==='A'?'selected':''}>Корпус A</option><option value="B" ${state.appealBuildingFilter==='B'?'selected':''}>Корпус B</option><option value="missing" ${state.appealBuildingFilter==='missing'?'selected':''}>Корпус не указан</option></select>`}<span class="list-count">${list.length}</span></div>
      <div class="appeal-list">${list.length?list.map(a=>appealRow(a,onlyClarify)).join(''):'<div class="empty-state"><strong>Ничего не найдено</strong>Измените фильтры или запрос</div>'}</div>`;
  }
  function worksView(){
    const q=norm(state.query);
    const works=sortedWorks(state.works.filter(w=>
      (!q||norm([w.id,w.location,w.description,w.building,w.zone,w.executor].join(' ')).includes(q))&&
      (state.workStatusFilter==='all'||norm(w.status)===norm(state.workStatusFilter))
    ));
    return `<div class="page-heading"><div><h1 class="page-title">Работы</h1></div><div class="heading-actions"><button class="btn btn-light" id="downloadCsvBtn">↓ Выгрузить список</button></div></div>
      <div class="table-toolbar"><div class="searchbox"><span>⌕</span><input id="workSearch" placeholder="Поиск по работе, комнате или исполнителю" value="${esc(state.query)}"></div>
      <select id="workFilter" class="filter-select"><option value="all" ${state.workStatusFilter==='all'?'selected':''}>Все статусы</option><option value="открыто" ${state.workStatusFilter==='открыто'?'selected':''}>Открыто</option><option value="принято" ${state.workStatusFilter==='принято'?'selected':''}>Принято</option><option value="в работе" ${state.workStatusFilter==='в работе'?'selected':''}>В работе</option><option value="закрыто" ${state.workStatusFilter==='закрыто'?'selected':''}>Закрыто</option></select>
      <select id="workSort" class="filter-select" aria-label="Сортировка работ"><option value="newest" ${state.workSort==='newest'?'selected':''}>Сначала обновлённые</option><option value="oldest" ${state.workSort==='oldest'?'selected':''}>Сначала давно обновлённые</option></select></div>
      <div class="appeal-list" id="workList">${works.map(w=>`<button type="button" class="problem-card work-feed-card" data-work="${esc(w.id)}"><span class="problem-main"><span class="problem-title">${esc(w.id)} · ${esc(w.location||w.rooms?.join(', ')||'Помещение не указано')} · ${esc(w.description)}</span><span class="problem-meta"><span>Корпус ${esc(w.building||'—')}</span><span>Зона ${esc(w.zone||'не указана')}</span><span>${w.appealIds.length} сообщ.</span><span>${esc(w.updated||'Время не указано')}</span></span></span><span class="status ${statusClass(w.status)}">${esc(statusLabel(w.status))}</span></button>`).join('')||'<div class="empty-state">Работы не найдены</div>'}</div>`;
  }
  function spacesView(){return `<div class="page-heading"><div><div class="eyebrow">СПРАВОЧНИК / ПОМЕЩЕНИЯ</div><h1 class="page-title">Помещения и зоны</h1><p class="page-subtitle">${state.spaces.length} записей · корпус — обязательная часть адреса</p></div></div><section class="panel"><table class="space-table"><thead><tr><th>Корпус</th><th>Помещение</th><th>Зона</th><th>Арендатор</th><th>Обслуживает</th></tr></thead><tbody>${state.spaces.map(s=>`<tr><td>${esc(s.building)}</td><td><strong>${esc(s.room)}</strong>${s.alias?`<div class="panel-hint">${esc(s.alias)}</div>`:''}</td><td>${esc(s.zone)}</td><td>${esc(s.tenant)}</td><td>${norm(s.owner).includes('арендатор')?'<span class="pill pill-purple">Арендатор</span>':'УК'}</td></tr>`).join('')}</tbody></table></section>`;}
  function rulesView(){return `<div class="page-heading"><div><div class="eyebrow">СПРАВОЧНИК / ПРАВИЛА</div><h1 class="page-title">Правила диспетчерской</h1><p class="page-subtitle">Из листа «Правила» исходной книги. Используются при группировке и уточнении.</p></div></div><section class="panel"><div class="panel-body">${state.rules.map(r=>`<div class="rule-row"><div class="rule-id">${esc(r.id)}</div><div class="rule-topic">${esc(r.topic)}</div><div class="rule-copy">${esc(r.text)}</div></div>`).join('')}</div></section>`;}
  function render(){
    const node=$('#view');
    if(!node)return;
    try{
      if(state.view==='overview')node.innerHTML=overviewView();
      else if(state.view==='appeals')node.innerHTML=appealsView(false);
      else if(state.view==='clarify')node.innerHTML=appealsView(true);
      else if(state.view==='resolved')node.innerHTML=resolvedProblemsView();
      else if(state.view==='spaces')node.innerHTML=spacesView();
      else if(state.view==='rules')node.innerHTML=rulesView();
      else node.innerHTML=overviewView();
      $$('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.view===state.view));
      if(state.selectedAppeal||state.selectedProblemWork)renderProblemDetail();
      if(state.selectedWork)openWorkDetail(state.selectedWork);
      if(state.selectedMessage)renderTenantMessageDetail();
    }catch(error){
      console.error('Ошибка отображения интерфейса:',error);
      node.innerHTML=`<section class="panel" style="max-width:900px;margin:36px auto;padding:28px;border:1px solid #f0c7c7"><h2 style="margin:0 0 10px;color:#a83232">Не удалось отобразить страницу</h2><p style="margin:0 0 12px;color:#536372">Произошла ошибка интерфейса. Перезагрузите приложение. Если ошибка повторится, отправьте технические сведения разработчику.</p><pre style="white-space:pre-wrap;overflow-wrap:anywhere;background:#fff5f5;padding:12px;border-radius:8px;color:#8f2c2c">${esc(error?.message||error)}</pre><button class="btn btn-primary" onclick="location.reload()">Перезагрузить приложение</button></section>`;
    }
  }
  function showStartupError(error){
    console.error('Ошибка запуска приложения:',error);
    const node=$('#view');
    if(node)node.innerHTML=`<section class="panel" style="max-width:900px;margin:36px auto;padding:28px;border:1px solid #f0c7c7"><h2 style="margin:0 0 10px;color:#a83232">Не удалось запустить Квартал</h2><p style="margin:0 0 12px;color:#536372">Проверьте рабочую книгу Excel в папке data и перезапустите приложение.</p><pre style="white-space:pre-wrap;overflow-wrap:anywhere;background:#fff5f5;padding:12px;border-radius:8px;color:#8f2c2c">${esc(error?.message||error)}</pre><button class="btn btn-primary" onclick="location.reload()">Повторить запуск</button></section>`;
    setSyncStatus('Ошибка запуска','sync-error');
  }
  function go(view){state.view=view;state.selectedAppeal=null;state.selectedProblemWork=null;state.selectedMessage=null;state.selectedWork=null;$('#overlayRoot').innerHTML='';state.query='';state.filter='all';if(view==='appeals'||view==='clarify'){state.appealChannelFilter='all';state.appealBuildingFilter='all';}render();}

  // Main problem card: not an individual tenant message. The linked messages
  // open as a second dialog above this one.
  function openAppeal(id){
    const a=state.appeals.find(x=>x.id===id);if(!a)return;
    state.selectedAppeal=id;state.selectedProblemWork=null;state.selectedMessage=null;state.selectedWork=null;renderProblemDetail();
  }
  function openProblemWork(id){
    const w=state.works.find(x=>x.id===id);if(!w)return;
    state.selectedAppeal=null;state.selectedProblemWork=id;state.selectedMessage=null;state.selectedWork=null;renderProblemDetail();
  }
  function closeAppealDetail(){state.selectedAppeal=null;state.selectedProblemWork=null;state.selectedMessage=null;state.selectedWork=null;$('#overlayRoot').innerHTML='';}
  function closeTenantMessage(){state.selectedMessage=null;const child=$('#tenantMessageOverlay');if(child)child.remove();}
  function workContext(w,group=null){
    const rawLocation=String(w.location||'').trim();
    const ids=(w.rooms&&w.rooms.length?w.rooms:roomList(rawLocation));
    const buildingHint=String(w.building||group?.building||'').toUpperCase();
    let space=null;
    if(ids.length)space=state.spaces.find(s=>norm(s.room)===norm(ids[0])&&(!buildingHint||s.building.toUpperCase()===buildingHint))||null;
    if(!space&&rawLocation)space=state.spaces.find(s=>norm(s.room)===norm(rawLocation)&&(!buildingHint||s.building.toUpperCase()===buildingHint))||null;
    // В старой книге в поле «Помещение или зона» иногда хранится именно зона.
    // В таком случае помещение берём из связанной проблемы, а зону — из справочника.
    const locationIsZone=!!rawLocation&&state.spaces.some(s=>norm(s.zone)===norm(rawLocation)&&(!buildingHint||s.building.toUpperCase()===buildingHint));
    const zoneSpace=locationIsZone?state.spaces.find(s=>norm(s.zone)===norm(rawLocation)&&(!buildingHint||s.building.toUpperCase()===buildingHint)):null;
    if(!space&&group?.room)space=state.spaces.find(s=>norm(s.room)===norm(group.room)&&(!buildingHint||s.building.toUpperCase()===buildingHint))||null;
    if(!space&&/холл/i.test(rawLocation))space=state.spaces.find(s=>s.building.toUpperCase()===buildingHint&&/холл/i.test(s.room))||null;
    const building=w.building||space?.building||zoneSpace?.building||group?.building||'';
    const rawRoom=ids.length?ids.map(x=>roomInputValue(x,building)).filter(Boolean).join(', '):(locationIsZone?'':roomInputValue(rawLocation,building));
    const room=rawRoom||roomInputValue(group?.room,building)||roomInputValue(space?.room,building)||'';
    const zone=w.zone||space?.zone||zoneSpace?.zone||group?.zone||'';
    return {building,room,zone,rawLocation};
  }

  function renderProblemDetail(){
    const id=state.selectedAppeal;
    const selectedStandalone=state.selectedProblemWork?state.works.find(x=>x.id===state.selectedProblemWork):null;
    const realAnchor=id?state.appeals.find(x=>x.id===id):null;
    const anchor=realAnchor||(selectedStandalone?{id:'',workId:selectedStandalone.id,building:selectedStandalone.building||'',room:selectedStandalone.location||'',zone:selectedStandalone.zone||'',issue:selectedStandalone.description||shortIssue(selectedStandalone.category||'прочее'),category:selectedStandalone.category||categoryOf(selectedStandalone.description||''),route:''}:null);
    if(!anchor){state.selectedAppeal=null;state.selectedProblemWork=null;$('#overlayRoot').innerHTML='';return;}
    const group=problemGroups().find(g=>selectedStandalone?g.work?.id===selectedStandalone.id:g.appeals.some(x=>x.id===id))||{appeals:realAnchor?[realAnchor]:[],latest:realAnchor,work:selectedStandalone||getAppealWork(anchor),issue:anchor.issue||shortIssue(anchor.category),building:appealBuilding(anchor),room:appealRoom(anchor),zone:anchor.zone||''};
    const w=group.work||getAppealWork(anchor)||null;
    const status=w?`<span class="status ${statusClass(w.status)}">${esc(statusLabel(w.status))}</span>`:'';
    const allQuestions=group.appeals.map(a=>({a,q:questionInfo(a)})).filter(x=>x.q);
    const openQuestions=allQuestions.filter(x=>!x.q.answered);
    const tenantMessages=sortedAppeals(group.appeals);
    const executorChoices=[...new Set(state.works.map(x=>String(x.executor||'').trim()).filter(x=>x&&norm(x)!=='не назначен'))].sort((a,b)=>a.localeCompare(b,'ru'));
    const executorList=`<datalist id="problemExecutorOptions">${executorChoices.map(x=>`<option value="${esc(x)}"></option>`).join('')}</datalist>`;
    const workDetails=w?`<div class="problem-work-details"><div class="problem-work-id">${esc(w.id)}</div><div class="problem-work-rows"><div><span>Обновлено</span><strong>${esc(w.updated||'—')}</strong></div></div><div class="problem-work-executor-edit"><label for="problemWorkExecutorInput">Исполнитель</label><div class="problem-work-executor-controls"><input id="problemWorkExecutorInput" list="problemExecutorOptions" value="${esc(norm(w.executor)==='не назначен'?'':w.executor||'')}" placeholder="Назначить исполнителя"><button type="button" class="btn btn-primary btn-small" id="saveProblemExecutorBtn">Сохранить</button></div>${executorList}</div><div class="problem-work-description-label">Описание работы</div><div class="problem-work-description">${esc(w.description||group.issue)}</div></div>`:anchor.route==='tenant'?'<div class="plain-note">Обслуживание помещения относится к зоне ответственности арендатора.</div>':'<div class="plain-note">Работа пока не заведена. Обращения сохранены, заявители не потеряны.</div>';
    const previousWorkId=w?.previousWork||'';const previousWork=previousWorkId?state.works.find(x=>x.id===previousWorkId):null;
    const previousWorkMarkup=previousWorkId?`<div class="work-history-item"><strong>№ предыдущей работы: ${esc(previousWorkId)}</strong><span>${esc(previousWork?.description||'Запись о предыдущей работе')}</span><small>Последнее обновление статуса: ${esc(previousWork?.updated||'Время в книге не указано')}</small></div>`:'';
    const workHistoryMarkup=`${previousWorkMarkup}${w?renderWorkHistory(w):'<div class="plain-note">История появится после создания работы.</div>'}`;
    const listMarkup=tenantMessages.map(x=>{
      const q=questionInfo(x);const badge=q&&!q.answered?'<span class="question-alert message-question-alert" title="Вопрос ожидает ответа">!</span>':'';
      return `<button type="button" class="problem-message-link" data-tenant-message="${esc(x.id)}"><span class="problem-message-copy"><strong>${badge}${esc(x.reporter||'Заявитель не указан')}</strong><span>${esc(x.text||'Сообщение без текста')}</span><small>${esc(x.time||'Время не указано')} · ${esc(x.channel||'Канал не указан')}</small></span><span class="problem-arrow">›</span></button>`;
    }).join('');
    $('#overlayRoot').innerHTML=`<div class="overlay appeal-detail-overlay" id="appealDetailOverlay"><section class="appeal-detail-modal problem-detail-modal" role="dialog" aria-modal="true" aria-labelledby="appealDetailTitle">
      <div class="appeal-modal-head"><div class="appeal-modal-title"><div class="eyebrow">ПРОБЛЕМА${group.appeals.length?` · ${group.appeals.length} ${appealCountWord(group.appeals.length)}`:' · создана диспетчером'}</div><div class="appeal-modal-id" id="appealDetailTitle">${esc(group.issue||'Неисправность')}</div><div class="detail-place-line">${esc([group.building?'Корпус '+group.building:'',group.room?roomInputValue(group.room,group.building):'',group.zone?'Зона обслуживания '+group.zone:''].filter(Boolean).join(' · ')||'Место уточняется')}</div></div><div class="appeal-modal-actions"><button type="button" class="close-btn appeal-detail-close" id="appealDetailClose" aria-label="Закрыть">×</button></div></div>
      <div class="appeal-detail-body"><div class="detail-workspace problem-workspace"><div class="detail-main">
        ${allQuestions.length?`<section class="detail-card pending-questions-card"><div class="detail-card-head"><span class="detail-card-title">Вопросы арендаторов</span>${openQuestions.length?'<span class="question-alert question-alert-large">!</span>':'<span class="question-answered-pill">На все вопросы ответили</span>'}</div>${allQuestions.map(({a,q})=>`<button type="button" class="pending-question-row ${q.answered?'is-answered':''}" data-tenant-message="${esc(a.id)}"><strong>${!q.answered?'<span class="question-alert message-question-alert">!</span>':''}${esc(a.reporter||'Заявитель')} ${q.answered?'<span class="question-row-answered">Ответ дан</span>':''}</strong><span>${esc(q.text)}</span><small>${esc(a.time||'')}</small></button>`).join('')}</section>`:''}
        <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Обращения арендаторов</span><span class="panel-hint">${tenantMessages.length}</span></div><div class="tenant-message-list">${listMarkup||(group.standalone?'<div class="empty-state">По этой проблеме пока нет сообщений арендаторов</div>':'<div class="empty-state">Нет связанных обращений</div>')}</div></section>
      </div><aside class="detail-side problem-work-side">
        <section class="detail-card problem-work-card"><div class="detail-card-head"><span class="detail-card-title">Работы</span>${w?status:''}</div>${workDetails}${!w&&anchor.route!=='tenant'&&group.building&&group.room?'<div class="problem-work-create"><button class="btn btn-primary" id="createWorkFromAppeal">＋ Создать работу</button></div>':''}</section>
        <section class="detail-card work-history-card"><div class="detail-card-head"><span class="detail-card-title">История работ</span></div><div class="work-history-list">${workHistoryMarkup}</div></section>
      </aside></div></div></section></div>`;
    $('#appealDetailClose').onclick=closeAppealDetail;
    $('#appealDetailOverlay').addEventListener('click',e=>{if(e.target.id==='appealDetailOverlay')closeAppealDetail();});
    if($('#saveProblemExecutorBtn'))$('#saveProblemExecutorBtn').onclick=()=>{
      const input=$('#problemWorkExecutorInput');
      const next=String(input?.value||'').trim();
      const previous=String(w.executor||'Не назначен').trim();
      const normalizedNext=norm(next);
      const assigned=!next||normalizedNext==='не назначен'?'Не назначен':next;
      if(norm(previous)===norm(assigned)){toast('Исполнитель не изменился');return;}
      w.executor=assigned;
      const previousStatus=w.status;
      if(assigned!=='Не назначен'&&['','открыто','без работы'].includes(norm(w.status)))w.status='принято';
      else if(assigned==='Не назначен'&&norm(w.status)==='принято')w.status='открыто';
      const statusChanged=norm(previousStatus)!==norm(w.status);
      if(statusChanged)w.updated=workStamp();
      const logText=`Исполнитель: ${previous} → ${assigned}.${statusChanged?' Статус автоматически изменён: '+previousStatus+' → '+w.status+'.':''}`;
      const historyAppeal=group.appeals[0]||null;
      if(historyAppeal)addHistory(historyAppeal.id,logText);
      addWorkHistory(w.id,logText,{status:w.status,statusUpdated:w.updated,kind:statusChanged?'status':'assignment'});
      queueSave();updateNav();render();
      toast(previousStatus!==w.status?`Исполнитель назначен · статус: ${statusLabel(w.status)}`:'Исполнитель обновлён');
      openAppeal(id);
    };
    if($('#createWorkFromAppeal'))$('#createWorkFromAppeal').onclick=()=>{const created=createManualWork(anchor);toast(`Создана работа ${created.id}`);render();openAppeal(id);};
  }
  function openTenantMessage(id){
    const a=state.appeals.find(x=>x.id===id);if(!a)return;
    state.selectedMessage=id;renderTenantMessageDetail();
  }
  function renderTenantMessageDetail(){
    const id=state.selectedMessage;const a=state.appeals.find(x=>x.id===id);if(!a){state.selectedMessage=null;return;}
    const w=getAppealWork(a);const q=questionInfo(a);const building=appealBuilding(a);const room=appealRoom(a);const zone=a.zone||w?.zone||state.spaces.find(s=>s.building===building&&norm(s.room)===norm(room))?.zone||'';
    const sourceFields=[['№',a.id],['Время',a.time],['Канал',a.channel],...(normalizeChannel(a.channel)==='портал'?[['Номер на портале',a.portal||'Не указан']]:[]),['Заявитель',a.reporter||'Не указан']];
    const contextFields=[['Корпус',building||'Не указан'],['Помещение',roomInputValue(room,building)||'Не указано'],['Зона обслуживания',zone||'Не указана'],['Работа',w?`${w.id} · ${w.description}`:'Не заведена'],['Статус работы',w?.status||'—']];
    const existing=$('#tenantMessageOverlay');if(existing)existing.remove();
    const child=document.createElement('div');child.className='overlay tenant-message-overlay';child.id='tenantMessageOverlay';child.innerHTML=`<section class="tenant-message-modal" role="dialog" aria-modal="true" aria-labelledby="tenantMessageTitle"><div class="appeal-modal-head"><div class="appeal-modal-title"><div class="eyebrow">СООБЩЕНИЕ АРЕНДАТОРА</div><div class="appeal-modal-id" id="tenantMessageTitle">${esc(a.id)}</div><div class="detail-place-line">${esc(a.reporter||'Заявитель не указан')}</div></div><div class="appeal-modal-actions"><button type="button" class="close-btn" id="tenantMessageClose" aria-label="Закрыть">×</button></div></div><div class="tenant-message-body">
      <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Данные из листа «Обращения»</span><button class="btn btn-light" id="editTenantMessageBtn">Изменить</button></div><div class="tenant-source-grid">${sourceFields.map(([label,value])=>`<div class="tenant-source-field"><span>${esc(label)}</span><strong>${esc(value||'—')}</strong></div>`).join('')}</div><div class="separator"></div><div class="tenant-raw-text-title">Текст обращения из Excel</div><div class="detail-message tenant-raw-text">${esc(a.text||'Текст обращения не заполнен')}</div></section>
      <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Контекст проблемы</span><button class="btn btn-primary btn-small" id="viewProblemFromMessage">Просмотреть проблему ↗</button></div><div class="tenant-source-grid">${contextFields.map(([label,value])=>`<div class="tenant-source-field"><span>${esc(label)}</span><strong>${esc(value||'—')}</strong></div>`).join('')}</div></section>
      ${q?`<section class="detail-card tenant-question-card ${q.answered?'is-answered':''}"><div class="detail-card-head"><span class="detail-card-title">Вопрос арендатора</span>${q.answered?'<span class="question-answered-pill">Ответ дан</span>':'<span class="question-alert">!</span>'}</div><div class="detail-message">${esc(q.text)}</div><div class="two-col-actions" style="margin-top:13px"><button class="btn ${q.answered?'btn-light':'btn-primary'}" id="toggleQuestionAnswered">${q.answered?'Вернуть в ожидающие':'Отметить как отвеченный'}</button></div></section>`:''}
      <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Анализ обращения</span><button class="btn btn-light" id="aiAnalyzeBtn">Разобрать обращение</button></div><div id="aiResult" class="plain-note">ИИ проверит содержание, выделит важные факты и подскажет, что нужно уточнить. Запрос отправляется в Google только после нажатия кнопки и при наличии согласия.</div></section>
      <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Ответ заявителю</span><button class="btn btn-primary" id="generateReplyBtn">✦ Подготовить</button></div><div class="form-field"><textarea id="replyDraftText" rows="8" placeholder="Здесь появится готовый ответ арендатору.">${esc(a.responseDraft||'')}</textarea><div class="form-note reply-generation-note" id="replyGenerationNote">Подготовьте ответ и проверьте адресата и факты перед отправкой.</div></div><div class="two-col-actions" style="margin-top:12px"><button class="btn btn-light" id="saveReplyDraftBtn">Сохранить</button><button class="btn btn-light" id="copyReplyDraftBtn">Копировать</button><button class="btn btn-primary" id="sendReplyBtn">Отправить</button></div><div class="form-note">Кнопка «Отправить» здесь отмечает вопрос как отвеченный; фактическая отправка письма не выполняется.</div></section>
      </div></section>`;
    $('#overlayRoot').appendChild(child);
    $('#tenantMessageClose').onclick=closeTenantMessage;
    child.addEventListener('click',e=>{if(e.target===child)closeTenantMessage();});
    if($('#editTenantMessageBtn'))$('#editTenantMessageBtn').onclick=()=>openAppealForm(a);
    if($('#viewProblemFromMessage'))$('#viewProblemFromMessage').onclick=()=>{go('overview');openAppeal(a.id);};
    if($('#toggleQuestionAnswered'))$('#toggleQuestionAnswered').onclick=()=>markTenantQuestionAnswered(id,!(q?.answered));
    $('#generateReplyBtn').onclick=()=>generateResidentReply(a);
    $('#aiAnalyzeBtn').onclick=()=>analyzeWithAI(a);
    $('#saveReplyDraftBtn').onclick=()=>{a.responseDraft=$('#replyDraftText').value.trim();state.appState.tenantMessages=state.appState.tenantMessages||{};state.appState.tenantMessages[id]={...(state.appState.tenantMessages[id]||{}),responseDraft:a.responseDraft};addHistory(id,'Текст ответа заявителю сохранён.');queueSave();toast('Ответ сохранён');};
    $('#copyReplyDraftBtn').onclick=async()=>{const text=$('#replyDraftText').value;if(!text){toast('Сначала подготовьте ответ.','warn');return;}try{await navigator.clipboard.writeText(text);toast('Текст ответа скопирован');}catch(_){$('#replyDraftText').select();document.execCommand('copy');toast('Текст ответа скопирован');}};
    $('#sendReplyBtn').onclick=()=>{const current=questionInfo(state.appeals.find(x=>x.id===id));if(current&&!current.answered){markTenantQuestionAnswered(id,true,'Вопрос отмечен как отвеченный. Внешняя отправка не выполняется.');}else{toast('Открытого вопроса нет. Фактическая отправка не подключена.','warn');}};
  }
  function closeOverlay(){
    const workEdit=$('#workEditOverlay');if(workEdit){workEdit.remove();return;}
    if(state.selectedMessage){closeTenantMessage();return;}
    const workDetail=$('#workDetailOverlay');if(workDetail){workDetail.remove();state.selectedWork=null;return;}
    if(state.selectedAppeal||state.selectedProblemWork){closeAppealDetail();return;}
    $('#overlayRoot').innerHTML='';
  }
  function closeFormModal(){if(state.selectedAppeal||state.selectedProblemWork){renderProblemDetail();if(state.selectedMessage)renderTenantMessageDetail();}else{$('#overlayRoot').innerHTML='';}}
  function workStamp(date=new Date()){
    return date.toLocaleString('ru-RU',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
  }
  function addHistory(appealId,text){const time=new Date().toLocaleTimeString('ru-RU',{hour:'2-digit',minute:'2-digit'});state.history.push({appealId,text,time});const a=state.appeals.find(x=>x.id===appealId);if(a)a.changeLog=[a.changeLog,time+' · '+text].filter(Boolean).join('\n');queueSave();}
  function addWorkHistory(workId,text,meta={}){
    const now=new Date();const time=now.toLocaleString('ru-RU',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
    state.appState.workHistory=state.appState.workHistory||{};state.appState.workHistory[workId]=state.appState.workHistory[workId]||[];
    state.appState.workHistory[workId].push({workId,text,time,at:now.toISOString(),status:meta.status||'',statusUpdated:meta.statusUpdated||'',kind:meta.kind||'change'});
    queueSave();
  }
  function renderWorkHistory(w){
    if(!w)return '<div class="plain-note">История работ недоступна.</div>';
    const entries=[...(state.appState.workHistory?.[w.id]||[])];
    const snapshot=`<div class="work-history-item work-history-current"><strong>№ работы: ${esc(w.id)}</strong><span>Статус: ${esc(statusLabel(w.status))}</span><small>Последнее обновление статуса: ${esc(w.updated||'Время в книге не указано')}</small></div>`;
    const rows=entries.slice().reverse().map(h=>{
      const timing=h.statusUpdated?`Последнее обновление статуса: ${esc(h.statusUpdated)}`:`Дата записи: ${esc(h.time||'Время не указано')}`;
      return `<div class="work-history-item"><strong>№ работы: ${esc(h.workId||w.id)}</strong><span>${esc(h.text||'Изменение работы')}</span><small>${h.status?'Статус: '+esc(statusLabel(h.status))+' · ':''}${timing}</small></div>`;
    }).join('');
    return snapshot+(rows?`<div class="work-history-subhead">Предыдущие изменения</div>${rows}`:'');
  }
  function moveAppeal(id,targetId){const a=state.appeals.find(x=>x.id===id);if(!a)return;const old=a.workId;state.works.forEach(w=>{w.appealIds=w.appealIds.filter(x=>x!==id);});if(targetId){const w=state.works.find(x=>x.id===targetId);if(!w){toast('Работа не найдена.','error');return;}w.appealIds.push(id);a.workId=targetId;a.manualUnlinked=false;a.route='';a.needsClarification=!!(a.missing||[]).length;a.linkReason='Перемещено диспетчером';addHistory(id,`Перенесено из ${old||'очереди'} в ${targetId}.`);}else{a.workId='';a.manualUnlinked=true;a.needsClarification=!!(a.missing||[]).length;addHistory(id,`Снята привязка к работе ${old||''}.`);}updateNav();queueSave();toast(targetId?`Входящее перенесено в ${targetId}`:'Обращение оставлено без работы');render();openAppeal(id);}
  function createManualWork(a){const nums=state.works.map(w=>Number((w.id.match(/(\d+)$/)||[])[1]||0));const no=Math.max(8,...nums)+1;const w={id:'РП-'+String(no).padStart(2,'0'),building:a.building,location:a.room,zone:a.zone||'—',executor:'Не назначен',status:'открыто',updated:workStamp(),description:a.issue||shortIssue(a.category),rooms:[a.room],category:a.category,appealIds:[a.id],source:'pilot',previousWork:'',notes:[]};state.works.push(w);a.workId=w.id;a.manualUnlinked=false;a.route='';a.needsClarification=!!(a.missing&&a.missing.length);addHistory(a.id,`Создана новая работа ${w.id}.`);addWorkHistory(w.id,'Работа создана из обращения.',{status:w.status,statusUpdated:w.updated,kind:'created'});updateNav();queueSave();return w;}
  function openProblemForm(){
    const overlay=document.createElement('div');overlay.className='overlay modal-overlay';overlay.id='newProblemOverlay';
    overlay.innerHTML=`<div class="modal form-modal new-problem-modal" role="dialog" aria-modal="true" aria-labelledby="newProblemTitle"><div class="modal-head"><div><div class="eyebrow">ДИСПЕТЧЕРСКАЯ</div><h3 id="newProblemTitle">Новая проблема</h3></div><button type="button" class="close-btn" id="newProblemClose" aria-label="Закрыть">×</button></div><form id="newProblemForm"><div class="modal-body"><div class="form-grid"><div class="form-field full"><label for="newProblemDescription">Что произошло</label><textarea id="newProblemDescription" name="description" required rows="3" placeholder="Например: холодно в помещении, протекает труба, не закрывается окно"></textarea></div><div class="form-field"><label for="newProblemBuilding">Корпус</label><select id="newProblemBuilding" name="building"><option value="">Не указан</option><option value="A">Корпус A</option><option value="B">Корпус B</option></select></div><div class="form-field"><label for="newProblemRoom">Помещение</label><input id="newProblemRoom" name="room" placeholder="301, переговорная, холл 1 этажа"></div><div class="form-field"><label for="newProblemZone">Зона обслуживания</label><input id="newProblemZone" name="zone" placeholder="Определится по помещению"></div><div class="form-field"><label for="newProblemExecutor">Исполнитель (необязательно)</label><input id="newProblemExecutor" name="executor" placeholder="Назначить исполнителя"></div></div><div class="form-note">Создаётся самостоятельная проблема и запись в листе «Работы». Сообщение арендатора не создаётся.</div></div><div class="modal-foot"><button type="button" class="btn btn-light" id="newProblemCancel">Отмена</button><button type="submit" class="btn btn-primary">Создать проблему</button></div></form></div>`;
    $('#overlayRoot').appendChild(overlay);
    const close=()=>overlay.remove();
    $('#newProblemClose').onclick=close;$('#newProblemCancel').onclick=close;
    overlay.addEventListener('click',e=>{if(e.target===overlay)close();});
    const fillZone=()=>{
      const building=$('#newProblemBuilding').value.trim().toUpperCase();
      const room=canonicalRoomValue($('#newProblemRoom').value,building);
      const space=state.spaces.find(sp=>norm(sp.room)===norm(room)&&(!building||sp.building.toUpperCase()===building));
      if(space&&space.zone)$('#newProblemZone').value=space.zone;
    };
    $('#newProblemBuilding').addEventListener('change',fillZone);
    $('#newProblemRoom').addEventListener('change',fillZone);
    $('#newProblemForm').addEventListener('submit',e=>{
      e.preventDefault();
      const form=Object.fromEntries(new FormData(e.currentTarget).entries());
      const description=String(form.description||'').trim();
      if(!description){toast('Укажи, что произошло.','warn');return;}
      const building=String(form.building||'').trim().toUpperCase();
      const location=canonicalRoomValue(String(form.room||'').trim(),building);
      const space=state.spaces.find(sp=>norm(sp.room)===norm(location)&&(!building||sp.building.toUpperCase()===building));
      const executor=String(form.executor||'').trim()||'Не назначен';
      const numbers=state.works.map(w=>Number((String(w.id||'').match(/(\d+)$/)||[])[1]||0));
      const next=Math.max(0,...numbers)+1;
      const id='РП-'+String(next).padStart(2,'0');
      const updated=new Date().toLocaleString('ru-RU',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'});
      const work={id,building,location,rooms:location?(roomList(location).length?roomList(location):[location]):[],zone:String(form.zone||'').trim()||space?.zone||'',executor,status:executor==='Не назначен'?'открыто':'принято',updated,description,category:categoryOf(description),appealIds:[],source:'manual-problem',priority:'Обычный',previousWork:'',notes:[]};
      state.works.push(work);
      addWorkHistory(id,'Проблема создана диспетчером.',{status:work.status,statusUpdated:work.updated,kind:'created'});
      close();updateNav();queueSave();render();openProblemWork(id);
      toast(`Создана проблема ${id}`);
    });
  }

  function openAppealForm(appeal=null){
    const a=appeal||{id:'',time:'16.11 12:30',channel:'портал',portal:'',reporter:'',text:'',building:'',room:''};
    const editing=!!appeal;
    const channel=normalizeChannel(a.channel);
    const currentProblem=appeal?problemGroups().find(g=>g.appeals.some(x=>x.id===a.id)):null;
    const problemKey=g=>'problem:'+(g.work?'work:'+g.work.id:'appeal:'+g.appeals[0].id);
    const selectedRelation=currentProblem?problemKey(currentProblem):'auto';
    const problemOptions=problemGroups().map(g=>{
      const key=problemKey(g);const place=[g.building?'корпус '+g.building:'',roomInputValue(g.room,g.building)].filter(Boolean).join(', ')||'место уточняется';
      return `<option value="${esc(key)}" ${selectedRelation===key?'selected':''}>${esc(g.issue)} · ${esc(place)} · ${g.appeals.length} сообщ.</option>`;
    }).join('');
    $('#overlayRoot').innerHTML=`<div class="overlay modal-overlay" id="formOverlay"><div class="modal form-modal"><div class="modal-head"><div><div class="eyebrow">${editing?'РЕДАКТИРОВАНИЕ':'НОВАЯ ПРОБЛЕМА'}</div><h3>${editing?'Изменить обращение '+esc(a.id):'Создать проблему'}</h3></div><button class="close-btn" id="formClose">×</button></div><form id="appealForm"><div class="modal-body"><div class="form-grid"><div class="form-field"><label>Заявитель</label><input name="reporter" value="${esc(a.reporter)}" placeholder="Имя или организация"></div><div class="form-field"><label>Канал</label><select name="channel" id="appealChannel">${channel&&!['портал','почта','стойка регистрации корпуса A','стойка регистрации корпуса B','общий телефон стойки регистрации'].includes(channel)?`<option value="${esc(channel)}" selected>${esc(channel)}</option>`:''}<option value="портал" ${channel==='портал'?'selected':''}>Портал</option><option value="почта" ${channel==='почта'?'selected':''}>Почта</option><option value="стойка регистрации корпуса A" ${channel==='стойка регистрации корпуса A'?'selected':''}>Стойка регистрации · корпус A</option><option value="стойка регистрации корпуса B" ${channel==='стойка регистрации корпуса B'?'selected':''}>Стойка регистрации · корпус B</option><option value="общий телефон стойки регистрации" ${channel==='общий телефон стойки регистрации'?'selected':''}>Общий телефон стойки регистрации</option></select></div><div class="form-field"><label>Корпус</label><select name="building"><option value="">Не указан</option><option value="A" ${a.building==='A'?'selected':''}>A</option><option value="B" ${a.building==='B'?'selected':''}>B</option></select></div><div class="form-field"><label>Кабинет / помещение</label><input name="room" value="${esc(roomInputValue(a.room,a.building))}" placeholder="301, переговорная, холл 1 этажа…"></div><div class="form-field" id="portalField" ${channel==='портал'?'':'hidden'}><label>Номер на портале</label><input name="portal" value="${esc(a.portal)}" placeholder="Например, P112"></div><div class="form-field"><label>Время</label><input name="time" value="${esc(a.time)}" placeholder="16.11 12:40"></div><div class="form-field full"><label>Текст обращения</label><textarea name="text" placeholder="Что случилось? Какие детали сообщил заявитель?">${esc(a.text)}</textarea></div><div class="form-field full relation-picker"><label>Связи с проблемой</label><select name="linkTarget" id="linkTarget"><option value="auto" ${selectedRelation==='auto'?'selected':''}>Добавить автоматически</option><option value="">Оставить без работы</option><optgroup label="Проблемы">${problemOptions||'<option disabled>Нет других проблем</option>'}</optgroup></select><span class="relation-help">Автоматически найдём подходящую открытую проблему или создадим новую. Выбранное обращение сохранится отдельным сообщением.</span></div></div><div class="form-note">Можно сохранить обращение с незаполненными полями.</div></div><div class="modal-foot"><button type="button" class="btn btn-light" id="cancelForm">Отмена</button><button type="submit" class="btn btn-primary">${editing?'Сохранить изменения':'Создать проблему'}</button></div></form></div></div>`;
    const togglePortalField=()=>{const visible=$('#appealChannel').value==='портал';$('#portalField').hidden=!visible;};
    $('#appealChannel').addEventListener('change',togglePortalField);togglePortalField();
    $('#formClose').onclick=closeFormModal;$('#cancelForm').onclick=closeFormModal;$('#formOverlay').addEventListener('click',e=>{if(e.target.id==='formOverlay')closeFormModal();});
    $('#appealForm').addEventListener('submit',e=>{e.preventDefault();const fd=new FormData(e.currentTarget);const obj=Object.fromEntries(fd.entries());obj.channel=normalizeChannel(obj.channel);obj.room=canonicalRoomValue(obj.room,obj.building);obj.portal=obj.channel==='портал'?(obj.portal||''):'';saveAppeal(a,obj,editing);});
  }
  function deriveAppeal(a){
    const text=a.text||'';let sp=null;
    if(a.room){sp=state.spaces.find(s=>norm(s.room)===norm(a.room)&&(!a.building||s.building.toUpperCase()===a.building.toUpperCase()));}
    if(!sp)sp=getSpaceForText((a.building?'корпус '+a.building+' ':'')+(a.room||'')+' '+text,a.reporter,state.spaces);
    const cat=categoryOf(text);
    const building=a.building||sp?.building||'';const room=a.room||sp?.room||'';
    const missing=[];if(!a.reporter||norm(a.reporter)==='не записан')missing.push('Заявитель / контакт');if(!building)missing.push('Корпус');if(!room)missing.push('Помещение');
    return {...a,building,room,zone:sp?.zone||'',owner:sp?.owner||'',category:cat,issue:shortIssue(cat),workId:a.workId||'',missing,needsClarification:missing.length>0,route:sp&&norm(sp.owner).includes('арендатор')?'tenant':''};
  }
  function saveAppeal(old,obj,editing){
    const a=deriveAppeal({...old,...obj});
    if(editing){Object.assign(old,a);}
    else{
      const max=state.appeals.reduce((m,x)=>Math.max(m,Number((x.id.match(/(\d+)$/)||[])[1]||0),0),0);
      a.id='О-'+String(max+1).padStart(2,'0');a.source='manual';a.time=a.time||'сейчас';
      state.appeals.push(a);
    }

    const relation=String(obj.linkTarget||'');
    const legacyAppealId=relation.startsWith('appeal:')?relation.slice(7):'';
    const targetProblemKey=relation.startsWith('problem:')?relation.slice(8):'';
    const targetProblem=targetProblemKey?problemGroups().find(g=>(g.work?'work:'+g.work.id:'appeal:'+g.appeals[0].id)===targetProblemKey):null;
    const target=targetProblem?(targetProblem.appeals.find(x=>x.id!==a.id)||targetProblem.appeals[0]):(legacyAppealId?state.appeals.find(x=>x.id===legacyAppealId):null);
    const targetWork=targetProblem?.work||(target?.workId?state.works.find(w=>w.id===target.workId):null);
    const autoAssign=relation==='auto'||(!relation&&!editing);
    // Сначала снимаем прежнюю привязку, затем добавляем выбранную проблему — так связи не дублируются.
    state.works.forEach(w=>{w.appealIds=w.appealIds.filter(x=>x!==a.id);});
    a.workId='';a.duplicateOf='';a.manualUnlinked=false;

    if(targetProblem?.work){
      const targetMessage=targetProblem.appeals.find(x=>x.id!==a.id)||targetProblem.appeals[0]||null;
      if(targetMessage)a.duplicateOf=targetMessage.id;
      a.linkReason='Связано диспетчером с проблемой '+targetProblem.work.id;
      if(!targetProblem.work.appealIds.includes(a.id))targetProblem.work.appealIds.push(a.id);
      a.workId=targetProblem.work.id;
    }else if(target){
      a.duplicateOf=target.id;a.linkReason='Связано диспетчером с проблемой '+(targetWork?.id||target.id);
      if(targetWork){if(!targetWork.appealIds.includes(a.id))targetWork.appealIds.push(a.id);a.workId=targetWork.id;}
      else if(target.building&&target.room&&target.route!=='tenant'){
        const sharedWork=createManualWork(target);
        if(!sharedWork.appealIds.includes(a.id))sharedWork.appealIds.push(a.id);
        a.workId=sharedWork.id;
      }else{a.manualUnlinked=true;if(target.route==='tenant')a.route='tenant';}
    }else if(autoAssign){
      if(a.route==='tenant'){a.routeText='Обслуживает арендатор по справочнику помещений';}
      else if(a.building&&a.room){
        const candidate=state.works.find(w=>norm(w.status)!=='закрыто'&&(!w.building||String(w.building).toUpperCase()===String(a.building).toUpperCase())&&(a.category==='температура'?a.zone&&w.zone===a.zone&&w.category==='температура':(w.rooms||[]).some(r=>norm(roomInputValue(r,w.building))===norm(roomInputValue(a.room,a.building)))&&w.category===a.category));
        if(candidate){candidate.appealIds.push(a.id);a.workId=candidate.id;a.linkReason='Автоматически добавлено к '+candidate.id;}
        else createManualWork(a);
      }else a.manualUnlinked=true;
    }else if(editing){
      a.manualUnlinked=true;a.linkReason='';
    }else{
      a.manualUnlinked=true;
    }

    if(editing){Object.assign(old,a);addHistory(old.id,'Данные обращения изменены.');}
    updateNav();if(target||targetProblem?.work){addHistory(a.id,'Добавлено в проблему '+(targetWork?.id||targetProblem?.work?.id||target?.id)+'.');}else if(autoAssign&&a.workId){addHistory(a.id,'Автоматически добавлено в проблему '+a.workId+'.');}
    queueSave();closeOverlay();toast(editing?'Изменения сохранены':'Обращение создано');
    render();openAppeal(a.id);
  }

  // Облачный ИИ. Модель запускается только на стороне Google после явного действия пользователя.
  // Никаких локальных моделей, WebGPU, Worker или фоновой загрузки.
  const GEMINI_CONSENT_KEY = 'kvartal_gemini_cloud_consent_v1';
  let geminiStatusCache = null;

  function hasGeminiConsent(){
    try{return localStorage.getItem(GEMINI_CONSENT_KEY)==='yes';}catch(_){return false;}
  }
  function setGeminiConsent(value){
    try{if(value)localStorage.setItem(GEMINI_CONSENT_KEY,'yes');else localStorage.removeItem(GEMINI_CONSENT_KEY);}catch(_){}
  }
  async function getGeminiStatus(force=false){
    if(geminiStatusCache&&!force)return geminiStatusCache;
    const response=await fetch('/api/ai/status',{cache:'no-store'});
    const data=await response.json();
    if(!response.ok)throw new Error(data.error||'Не удалось проверить настройки Gemini');
    geminiStatusCache=data;return data;
  }
  async function cloudAIRequest(task,context){
    if(!hasGeminiConsent())throw new Error('Сначала подтвердите передачу обезличенного текста в Google в настройках помощника.');
    const response=await fetch('/api/ai/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({task,context,consent:true})});
    const data=await response.json();
    if(!response.ok)throw new Error(data.error||'Не удалось получить ответ Gemini');
    return data.result||{};
  }
  function appealAIContext(a){
    const w=getAppealWork(a);
    return {
      id:a.id, issue:a.issue||shortIssue(a.category||'прочее'), text:a.text||'', category:a.category||'',
      building:a.building||'', room:a.room||'', zone:a.zone||'', route:a.route||'',
      missing:(a.missing||[]), question:questionInfo(a)?.text||'', time:a.time||'', channel:a.channel||'',
      work_id:w?.id||'', work_status:w?.status||'', work_description:w?.description||'',
      executor_assigned:!!(w?.executor&&norm(w.executor)!=='не назначен')
    };
  }
  function fallbackAnalysis(a){
    const missing=[...(a.missing||[])].filter(Boolean);
    const important=[];
    const place=[a.building?'корпус '+a.building:'',a.room].filter(Boolean).join(', ');
    if(place)important.push('Место: '+place);
    if(a.channel)important.push('Канал: '+a.channel);
    const work=getAppealWork(a);
    if(work)important.push('Связанная работа: '+work.id+' · '+work.status);
    return {summary:a.issue||shortIssue(a.category||'прочее')||'Требуется проверить обращение',important:important.slice(0,4),missing,question:missing.length?'Уточните, пожалуйста: '+missing.join(', ').toLowerCase()+'.':''};
  }
  function makeResidentReply(a){
    const w=getAppealWork(a);
    const place=[a.building?'корпус '+a.building:'',a.room||''].filter(Boolean).join(', ')||'место требуется уточнить';
    const subject=`Тема: Обращение ${a.id} — ${place}`;
    let stateText='Обращение зарегистрировано и находится на рассмотрении диспетчерской службы.';
    let next='Диспетчерская служба проверит сведения и определит дальнейший порядок передачи обращения ответственному исполнителю.';
    if(a.route==='tenant'){
      stateText='По справочнику помещений обслуживание указанного оборудования/помещения относится к зоне ответственности арендатора. Новая работа управляющей компании по этому обращению не назначалась.';
      next='Рекомендуем передать сообщение обслуживающей организации арендатора. При необходимости диспетчерская служба может уточнить сведения по границам ответственности.';
    }else if(w){
      stateText=`Обращение связано с работой ${w.id}. Текущий статус работы в системе: «${w.status}». Исполнитель ${w.executor&&norm(w.executor)!=='не назначен'?'назначен':'не назначен'}.`;
      if(norm(w.status)==='закрыто')next='В системе работа отмечена как закрытая. Если неисправность сохраняется или возникла повторно, просим сообщить, когда именно проблема была замечена снова.';
      else if(norm(w.status)==='принято')next='Работа принята исполнителем; после обновления статуса диспетчерская служба сможет сообщить дополнительную информацию.';
      else if(norm(w.status)==='в работе')next='Работа находится в выполнении. После получения отметки исполнителя о результате диспетчерская служба направит обновление.';
      else next='Работа зарегистрирована; диспетчерская служба контролирует её дальнейшее назначение и обновление статуса.';
    }
    const missing=(a.missing||[]).filter(Boolean);
    const clarification=missing.length?`\n\nДля корректной обработки просим уточнить: ${missing.join(', ').toLowerCase()}. До получения этих сведений обращение сохраняется в очереди и не теряется.`:'';
    return `${subject}\n\nЗдравствуйте!\n\nБлагодарим за обращение в управляющую компанию «Квартал Менеджмент». Подтверждаем, что сообщение зарегистрировано в диспетчерской системе под номером ${a.id}${a.time?' (поступило '+a.time+')':''}.\n\nСуть обращения: ${a.text||a.issue||'требуется уточнение описания неисправности'}.\nМесто: ${place}.\n${stateText}\n\n${next}${clarification}\n\nЕсли у вас есть дополнительные сведения, фотографии или уточнения по месту и времени возникновения неисправности, пожалуйста, направьте их ответным сообщением с указанием номера ${a.id}.\n\nС уважением,\nДиспетчерская служба\nУправляющая компания «Квартал Менеджмент»`;
  }
  async function generateResidentReply(a){
    const btn=$('#generateReplyBtn');const area=$('#replyDraftText');
    if(!btn||!area)return;
    btn.disabled=true;btn.textContent='Готовим текст…';
    let answer=makeResidentReply(a);let source='встроенному шаблону';
    try{
      const status=await getGeminiStatus(true);
      if(status.configured&&hasGeminiConsent()){
        const result=await cloudAIRequest('reply',appealAIContext(a));
        const candidate=String(result.text||'').trim();
        if(candidate.length<40)throw new Error('Gemini вернула слишком короткий черновик. Использован безопасный шаблон.');
        answer=candidate;source='Gemini';
      }
    }catch(error){
      console.warn('[Квартал] Gemini unavailable; template fallback used:',error?.message||error);
      source='резервному шаблону';
    }
    a.responseDraft=answer;area.value=answer;
    state.appState.tenantMessages=state.appState.tenantMessages||{};
    state.appState.tenantMessages[a.id]={...(state.appState.tenantMessages[a.id]||{}),responseDraft:answer};
    queueSave();addHistory(a.id,'Подготовлен черновик ответа по '+source+'.');
    const note=$('#replyGenerationNote');
    if(note){
      note.textContent=source==='Gemini'?'Черновик подготовлен Gemini по обезличенным данным обращения. Проверьте адресата и факты перед отправкой.':`Использован ${source}: Gemini не настроена, согласие не дано или сервис временно недоступен. Компьютер не запускает локальную модель.`;
    }
    btn.disabled=false;btn.textContent='Подготовить заново';
  }
  async function openAISettings(){
    $('#overlayRoot').innerHTML=`<div class="overlay modal-overlay" id="aiSettingsOverlay"><div class="modal" style="max-width:560px"><div class="modal-head"><div><div class="eyebrow">ОБЛАЧНЫЙ ИИ · БЕЗ НАГРУЗКИ НА ПК</div><h3>Gemini Flash-Lite</h3></div><button class="close-btn" id="aiClose">×</button></div><div class="modal-body">
      <div class="source-file"><div class="file-icon">G</div><div class="file-meta"><strong>Google Gemini 3.1 Flash-Lite</strong><span>Работает на серверах Google · компьютер не запускает модель</span></div><span class="ai-model-pill">Cloud API</span></div>
      <div class="form-note" style="margin-top:14px">Для бесплатного доступа создай API-ключ в Google AI Studio. Проект использует только модель gemini-3.1-flash-lite, без платных резервных моделей. Запрос отправляется только после нажатия кнопки ИИ.</div>
      <div class="form-field" style="margin-top:14px"><label for="geminiApiKey">API-ключ Google AI Studio</label><input id="geminiApiKey" type="password" autocomplete="new-password" spellcheck="false" placeholder="Вставь ключ; сохранённый ключ не показывается"><div class="form-note">Ключ хранится в личной папке пользователя на сервере, вне папки сайта; в браузер и Git не попадает.</div></div>
      <label class="check-row" style="display:flex;gap:10px;align-items:flex-start;margin-top:12px"><input id="geminiPrivacyConsent" type="checkbox" ${hasGeminiConsent()?'checked':''} style="margin-top:3px"><span>Я согласен отправлять Google только необходимые обезличенные фрагменты обращений для анализа и подготовки ответов. Понимаю, что на бесплатном тарифе Google может использовать отправленный контент для улучшения продуктов. Не отправлять специальные категории персональных данных и конфиденциальные сведения.</span></label>
      <div id="geminiStatusBox" class="soft-note" style="margin-top:14px">Проверяем настройки…</div>
      <div class="form-note" style="margin-top:10px">Перед отправкой сервер скрывает email, телефон и некоторые явные упоминания имени. Но автоматическое обезличивание не гарантирует удаление всех персональных данных: проверь текст обращения самостоятельно. Если бесплатная квота закончится, ИИ-запрос остановится и будет использован шаблон.</div>
    </div><div class="modal-foot" style="flex-wrap:wrap"><button class="btn btn-light" id="aiCancel">Закрыть</button><button class="btn btn-light" id="clearGeminiKey">Удалить ключ</button><button class="btn btn-light" id="testGeminiKey">Проверить ключ</button><button class="btn btn-primary" id="saveGeminiKey">Сохранить ключ</button></div></div></div>`;
    const statusBox=$('#geminiStatusBox');
    const refreshStatus=async()=>{
      try{const status=await getGeminiStatus(true);statusBox.textContent=status.configured?`Ключ настроен · модель ${status.model}. Согласие на отправку ${hasGeminiConsent()?'дано':'пока не дано'}.`:'API-ключ ещё не настроен. Пока будут использоваться шаблоны.';statusBox.className='soft-note '+(status.configured?'success-text':'');}
      catch(e){statusBox.textContent='Не удалось прочитать настройки: '+e.message;statusBox.className='soft-note';}
    };
    $('#aiClose').onclick=closeOverlay;$('#aiCancel').onclick=closeOverlay;
    $('#aiSettingsOverlay').addEventListener('click',e=>{if(e.target.id==='aiSettingsOverlay')closeOverlay();});
    $('#geminiPrivacyConsent').onchange=e=>{setGeminiConsent(!!e.target.checked);};
    $('#saveGeminiKey').onclick=async()=>{
      const key=$('#geminiApiKey').value.trim();
      if(!key){statusBox.textContent='Вставь новый API-ключ. Существующий ключ не отображается.';return;}
      const btn=$('#saveGeminiKey');btn.disabled=true;
      try{const r=await fetch('/api/ai/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({apiKey:key})});const d=await r.json();if(!r.ok)throw new Error(d.error||'Не удалось сохранить ключ');$('#geminiApiKey').value='';geminiStatusCache=null;await refreshStatus();statusBox.textContent='Ключ сохранён на сервере. Теперь нажми «Проверить ключ».';toast('API-ключ сохранён');}
      catch(e){statusBox.textContent=e.message;}
      finally{btn.disabled=false;}
    };
    $('#testGeminiKey').onclick=async()=>{
      const btn=$('#testGeminiKey');btn.disabled=true;
      try{
        setGeminiConsent($('#geminiPrivacyConsent').checked);
        const typed=$('#geminiApiKey').value.trim();
        if(typed){
          const save=await fetch('/api/ai/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({apiKey:typed})});
          const saved=await save.json();if(!save.ok)throw new Error(saved.error||'Не удалось сохранить ключ');
          $('#geminiApiKey').value='';geminiStatusCache=null;
        }
        statusBox.textContent='Отправляем один короткий тестовый запрос в Google…';
        const r=await fetch('/api/ai/test',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
        const d=await r.json();if(!r.ok)throw new Error(d.error||'Проверка не пройдена');
        await refreshStatus();statusBox.textContent=`Подключение успешно · ${d.model}. ${d.message||''}`;statusBox.className='soft-note success-text';toast('Gemini подключена');
      }catch(e){statusBox.textContent=e.message;statusBox.className='soft-note';}
      finally{btn.disabled=false;}
    };
    $('#clearGeminiKey').onclick=async()=>{
      if(!confirm('Удалить сохранённый API-ключ Gemini с этого сервера?'))return;
      try{const r=await fetch('/api/ai/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'clear'})});const d=await r.json();if(!r.ok)throw new Error(d.error||'Не удалось удалить ключ');geminiStatusCache=null;await refreshStatus();toast('Ключ удалён');}
      catch(e){statusBox.textContent=e.message;}
    };
    await refreshStatus();
  }
  async function analyzeWithAI(a){
    const result=$('#aiResult');const btn=$('#aiAnalyzeBtn');
    if(btn){btn.disabled=true;btn.textContent='Анализируем…';}
    let parsed=fallbackAnalysis(a);let source='по правилам проекта';
    try{
      const status=await getGeminiStatus(true);
      if(status.configured&&hasGeminiConsent()){
        const response=await cloudAIRequest('analysis',appealAIContext(a));
        parsed={summary:String(response.summary||parsed.summary),important:Array.isArray(response.important)?response.important.slice(0,5):parsed.important,missing:Array.isArray(response.missing)?response.missing.slice(0,8):parsed.missing,question:String(response.question||'')};source='Gemini';
      }
    }catch(error){console.warn('[Квартал] Analysis fallback:',error?.message||error);source='резервным правилам (Gemini недоступна)';}
    a.aiSummary=parsed.summary;state.aiAnalysis=state.aiAnalysis||{};state.aiAnalysis[a.id]=parsed;queueSave();
    if(result){
      result.innerHTML=`<div class="detail-message"><strong>${esc(parsed.summary||'')}</strong>${parsed.important.length?`<ul style="margin:8px 0 0;padding-left:18px">${parsed.important.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`:''}</div>${parsed.missing.length?`<div class="plain-note" style="margin-top:10px"><strong>Нужно уточнить:</strong> ${parsed.missing.map(esc).join(' · ')}${parsed.question?`<div style="margin-top:6px">${esc(parsed.question)}</div>`:''}</div>`:''}<div class="form-note">Результат получен ${source}. Проверьте вывод перед использованием.</div>`;
    }
    if(btn){btn.disabled=false;btn.textContent='Разобрать обращение';}
  }
  function makeBriefFallback(group,records){
    if(!group)return 'В книге пока нет зарегистрированных работ. Новые обращения появятся в разделе «Входящие» после загрузки или добавления данных.';
    const places=(group.rooms||[]).map(r=>roomInputValue(r,group.building)).filter(Boolean);
    const clarification=records.filter(a=>a.needsClarification);
    const lines=[`Работа: ${group.id}.`,`Проблема: ${group.description||'описание не указано'}.`,`Место: ${[group.building?'корпус '+group.building:'',places.join(', ')||group.location||'помещение не указано'].filter(Boolean).join(', ')}.`,`Зона обслуживания: ${group.zone||'не указана'}.`,`Статус: ${group.status||'не указан'}.`,`Исполнитель: ${group.executor&&norm(group.executor)!=='не назначен'?'назначен':'не назначен'}.`,`Связано обращений: ${records.length}.`];
    if(clarification.length)lines.push(`Требуют уточнения: ${clarification.map(a=>a.id).join(', ')}.`);
    if(records.length){const descriptions=[...new Set(records.map(a=>String(a.text||'').trim()).filter(Boolean))].slice(0,3);if(descriptions.length)lines.push('Сообщения: '+descriptions.join(' / '));}
    return lines.join(' ');
  }
  async function generateBrief(workId=null){
    const btn=$('#aiBriefBtn');if(btn){btn.disabled=true;btn.textContent='Формируем…';}
    const group=state.works.find(w=>w.id===workId)||state.works.find(w=>w.id==='Р-01')||[...state.works].sort((a,b)=>(b.appealIds||[]).length-(a.appealIds||[]).length)[0]||null;
    const records=group?state.appeals.filter(a=>(group.appealIds||[]).includes(a.id)):[];
    let summary=makeBriefFallback(group,records);let source='по правилам проекта';
    try{
      const status=await getGeminiStatus(true);
      if(status.configured&&hasGeminiConsent()&&group){
        const context={id:group.id,issue:group.category||'',work_description:group.description||'',building:group.building||'',room:group.location||((group.rooms||[]).join(', ')),zone:group.zone||'',work_status:group.status||'',executor_assigned:!!(group.executor&&norm(group.executor)!=='не назначен'),appeal_count:records.length,appeal_ids:records.map(a=>a.id),clarification_ids:records.filter(a=>a.needsClarification).map(a=>a.id),descriptions:records.map(a=>a.text||'').filter(Boolean).slice(0,8),previous_work:group.previousWork||''};
        const response=await cloudAIRequest('brief',context);if(String(response.text||'').trim().length>40){summary=String(response.text).trim();source='Gemini';}
      }
    }catch(error){console.warn('[Квартал] Brief fallback:',error?.message||error);source='резервным правилам (Gemini недоступна)';}
    if(group){group.aiSummary=summary;queueSave();}
    showSummaryModal(summary,'Сводка по работе '+(group?.id||''),`Сформировано ${source}. Проверьте статусы и факты перед использованием.`);
    if(btn){btn.disabled=false;btn.textContent='Сформировать сводку';}
  }

  function showSummaryModal(text,title,note='Сформировано по данным книги Excel. Проверьте факты перед использованием.'){
    $('#overlayRoot').innerHTML=`<div class="overlay modal-overlay" id="summaryOverlay"><div class="modal" style="max-width:480px"><div class="modal-head"><div><div class="eyebrow">КРАТКО И ПО ДЕЛУ</div><h3>${esc(title)}</h3></div><button class="close-btn" id="summaryClose">×</button></div><div class="modal-body"><div class="detail-message">${esc(text)}</div><div class="form-note">${esc(note)}</div></div><div class="modal-foot"><button class="btn btn-light" id="summaryClose2">Закрыть</button><button class="btn btn-primary" id="copySummary">Копировать</button></div></div></div>`;
    $('#summaryClose').onclick=closeOverlay;$('#summaryClose2').onclick=closeOverlay;
    $('#summaryOverlay').addEventListener('click',e=>{if(e.target.id==='summaryOverlay')closeOverlay();});
    $('#copySummary').onclick=()=>{navigator.clipboard?.writeText(text);toast('Сводка скопирована');};
  }

  function downloadCSV(){const rows=[['Работа','Корпус','Помещение','Статус','Исполнитель','Обращения'],...state.works.map(w=>[w.id,w.building,w.location,w.status,w.executor,w.appealIds.join(', ')])];const csv='\uFEFF'+rows.map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(';')).join('\r\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='Квартал_работы.csv';a.click();URL.revokeObjectURL(url);}
  async function replaceFile(file){if(!file)return;if(!file.name.toLowerCase().endsWith('.xlsx')){toast('Нужен файл Excel .xlsx','error');return;}if(file.size>20*1024*1024){toast('Файл больше 20 МБ','error');return;}toast('Читаю книгу Excel…');try{const response=await fetch('/api/import',{method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','X-Filename':encodeURIComponent(file.name)},body:file});const data=await response.json();if(!response.ok)throw new Error(data.error||'Не удалось прочитать файл');loadData(data.sheets,file.name,data.revision,data.appState,data.appRevision);setSyncStatus('Excel синхронизирован','sync-ok');toast(`Загружен файл ${file.name}`);}catch(e){toast(e.message,'error');}}

  document.addEventListener('click',e=>{
    const nav=e.target.closest('[data-view]');if(nav){go(nav.dataset.view);return;}
    const goBtn=e.target.closest('[data-go]');if(goBtn){go(goBtn.dataset.go);return;}
    const kpi=e.target.closest('[data-kpi-filter]');if(kpi){state.filter=state.filter===kpi.dataset.kpiFilter?'all':kpi.dataset.kpiFilter;render();return;}
    const tenantMessage=e.target.closest('[data-tenant-message]');if(tenantMessage){openTenantMessage(tenantMessage.dataset.tenantMessage);return;}
    const problemWork=e.target.closest('[data-problem-work]');if(problemWork){openProblemWork(problemWork.dataset.problemWork);return;}
    const appeal=e.target.closest('[data-appeal]');if(appeal){openAppeal(appeal.dataset.appeal);return;}
    const work=e.target.closest('[data-work]');if(work){const w=state.works.find(x=>x.id===work.dataset.work);if(w)openWork(w.id);return;}
    if(e.target.closest('#quickNewBtn')){openProblemForm();return;}
    if(e.target.closest('#aiSettingsBtn')){openAISettings();return;}
    if(e.target.closest('#aiBriefBtn')){generateBrief();return;}
    if(e.target.closest('#showExampleBtn')){const w=state.works.find(x=>x.id==='Р-01');if(w)openWork(w.id);return;}
    if(e.target.closest('#downloadCsvBtn')){downloadCSV();return;}
  });
  document.addEventListener('keydown',e=>{if(e.key==='Escape')closeOverlay();});
  document.addEventListener('input',e=>{
    if(e.target.id==='appealSearch'){state.query=e.target.value;const pos=e.target.selectionStart;render();const el=$('#appealSearch');el?.focus();el?.setSelectionRange(pos,pos);}
    if(e.target.id==='workSearch'){state.query=e.target.value;const pos=e.target.selectionStart;render();const el=$('#workSearch');el?.focus();el?.setSelectionRange(pos,pos);}
  });
  document.addEventListener('change',e=>{
    if(e.target.id==='replaceFile'){replaceFile(e.target.files?.[0]);e.target.value='';}
    if(e.target.id==='appealFilter'){state.filter=e.target.value;render();}
    if(e.target.id==='appealChannelFilter'){state.appealChannelFilter=e.target.value;render();}
    if(e.target.id==='appealBuildingFilter'){state.appealBuildingFilter=e.target.value;render();}
    if(e.target.id==='overviewBuildingFilter'){state.overviewBuildingFilter=e.target.value;render();}
    if(e.target.id==='workFilter'){state.workStatusFilter=e.target.value;render();}
    if(e.target.id==='workSort'){state.workSort=e.target.value;render();}
  });

  function openWork(id){openWorkDetail(id);}
  function openWorkDetail(id){
    const w=state.works.find(x=>x.id===id);if(!w)return;
    state.selectedWork=id;
    const group=problemGroups().find(g=>g.work?.id===id)||null;
    const ctx=workContext(w,group);
    const linked=state.appeals.filter(a=>w.appealIds.includes(a.id));
    const historyMarkup=renderWorkHistory(w);
    const linkedMarkup=linked.map(a=>`<button class="work-linked-message" type="button" data-tenant-message="${esc(a.id)}"><strong>${esc(a.reporter||'Заявитель не указан')}</strong><span>${esc(a.text||'')}</span><small>${esc(a.time||'')} · ${esc(a.channel||'Канал не указан')}</small></button>`).join('');
    const old=$('#workDetailOverlay');if(old)old.remove();
    const node=document.createElement('div');node.className='overlay work-detail-overlay';node.id='workDetailOverlay';
    node.innerHTML=`<section class="work-detail-modal" role="dialog" aria-modal="true" aria-labelledby="workDetailTitle">
      <div class="appeal-modal-head"><div class="appeal-modal-title"><div class="eyebrow">КАРТОЧКА РАБОТЫ</div><div class="appeal-modal-id" id="workDetailTitle">${esc(w.id)}</div><div class="detail-place-line">${esc(ctx.building?'Корпус '+ctx.building:'Корпус не указан')} · ${esc(ctx.room||'Помещение не указано')}</div></div><div class="appeal-modal-actions"><span class="status ${statusClass(w.status)}">${esc(statusLabel(w.status))}</span><button type="button" class="close-btn" id="workDetailClose" aria-label="Закрыть">×</button></div></div>
      <div class="work-detail-body">
        <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Данные работы</span><button class="btn btn-light btn-small" id="aiBriefBtn">Сводка ИИ</button></div>
          <div class="tenant-source-grid work-detail-fields"><div class="tenant-source-field"><span>Корпус</span><strong>${esc(ctx.building||'Не указан')}</strong></div><div class="tenant-source-field"><span>Помещение</span><strong>${esc(ctx.room||'Не указано')}</strong></div><div class="tenant-source-field"><span>Зона обслуживания</span><strong>${esc(ctx.zone||'Не указана')}</strong></div><div class="tenant-source-field work-detail-value-field"><span>Обновление</span><strong>${esc(w.updated||'—')}</strong></div><div class="tenant-source-field"><span>Исполнитель</span><strong>${esc(w.executor||'Не назначен')}</strong></div><div class="tenant-source-field"><span>Связанные обращения</span><strong>${linked.length}</strong></div><div class="tenant-source-field work-detail-description-field"><span>Описание работы</span><strong>${esc(w.description||'Описание не заполнено')}</strong></div></div>
          ${w.previousWork?`<div class="missing-banner" style="margin-top:12px">Связана с предыдущей работой ${esc(w.previousWork)}.</div>`:''}
        </section>
        <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">Обращения арендаторов</span><span class="panel-hint">${linked.length}</span></div><div class="work-linked-list">${linkedMarkup||'<div class="empty-state">Нет связанных обращений</div>'}</div></section>
        <section class="detail-card"><div class="detail-card-head"><span class="detail-card-title">История работы</span></div><div class="work-history-list">${historyMarkup}</div></section>
      </div>
    </section>`;
    $('#overlayRoot').appendChild(node);
    $('#workDetailClose').onclick=()=>{node.remove();state.selectedWork=null;};
    node.addEventListener('click',e=>{if(e.target===node){node.remove();state.selectedWork=null;}});
  }
  function editWork(w){
    const ctx=workContext(w,problemGroups().find(g=>g.work?.id===w.id)||null);
    const node=document.createElement('div');node.className='overlay modal-overlay work-edit-overlay';node.id='workEditOverlay';
    node.innerHTML=`<div class="modal work-edit-modal" role="dialog" aria-modal="true"><div class="modal-head"><div><div class="eyebrow">РАБОТА ${esc(w.id)}</div><h3>Изменить работу</h3></div><button class="close-btn" id="workEditClose" aria-label="Закрыть">×</button></div><div class="modal-body">
      <div class="work-edit-summary"><strong>${esc(w.id)}</strong><span>${esc(ctx.building?'Корпус '+ctx.building:'Корпус не указан')} · ${esc(ctx.room||'Помещение не указано')} · ${esc(ctx.zone||'Зона не указана')}</span></div>
      <div class="form-grid"><div class="form-field"><label>Корпус</label><select id="workBuilding"><option value="">Не указан</option><option value="A" ${ctx.building==='A'?'selected':''}>A</option><option value="B" ${ctx.building==='B'?'selected':''}>B</option></select></div><div class="form-field"><label>Помещение</label><input id="workLocation" value="${esc(ctx.rawLocation||ctx.room)}"></div><div class="form-field"><label>Зона обслуживания</label><input id="workZone" value="${esc(ctx.zone)}"></div><div class="form-field"><label>Статус работы</label><select id="workStatus">${['открыто','принято','в работе','закрыто'].map(s=>`<option value="${s}" ${norm(s)===norm(w.status)?'selected':''}>${s}</option>`).join('')}</select></div><div class="form-field full"><label>Исполнитель</label><input id="workExecutor" value="${esc(w.executor||'')}"></div><div class="form-field full work-description-field"><label for="workDescription">Описание работы</label><textarea id="workDescription" class="work-description-editor" rows="5" spellcheck="true">${esc(w.description||'')}</textarea></div></div>
      <div class="form-note">Изменения сохранятся в рабочей книге Excel.</div></div><div class="modal-foot"><button class="btn btn-light" id="workEditCancel">Отмена</button><button class="btn btn-primary" id="workEditSave">Сохранить изменения</button></div></div>`;
    $('#overlayRoot').appendChild(node);
    const closeEdit=()=>node.remove();
    $('#workEditClose').onclick=closeEdit;$('#workEditCancel').onclick=closeEdit;
    node.addEventListener('click',e=>{if(e.target===node)closeEdit();});
    $('#workEditSave').onclick=()=>{
      const prevStatus=w.status;const prevDescription=w.description;const prevExecutor=w.executor;w.building=$('#workBuilding').value.trim().toUpperCase();w.location=(roomList($('#workLocation').value).length?roomList($('#workLocation').value).join(', '):canonicalRoomValue($('#workLocation').value,w.building));w.rooms=roomList(w.location);if(!w.rooms.length&&w.location){const matchedSpace=state.spaces.find(sp=>norm(sp.room)===norm(w.location)&&(!w.building||sp.building.toUpperCase()===w.building));w.rooms=[matchedSpace?.room||w.location];}w.zone=$('#workZone').value.trim();w.status=$('#workStatus').value;w.executor=$('#workExecutor').value.trim();w.description=$('#workDescription').value.trim();
      const statusChanged=norm(prevStatus)!==norm(w.status);
      if(statusChanged)w.updated=workStamp();
      state.spaces.forEach(sp=>{if(norm(sp.room)===norm(w.location)&&(!w.building||sp.building.toUpperCase()===w.building))w.zone=w.zone||sp.zone;});
      const changed=[];if(statusChanged)changed.push(`статус: ${prevStatus||'не указан'} → ${w.status}`);if(prevExecutor!==w.executor)changed.push(`исполнитель: ${prevExecutor||'не назначен'} → ${w.executor||'не назначен'}`);if(prevDescription!==w.description)changed.push('обновлено описание работы');if(!changed.length)changed.push('проверены данные работы');
      const historyText=`Работа ${w.id}: ${changed.join('; ')}.`;
      const linkedForHistory=state.appeals.filter(a=>w.appealIds.includes(a.id));
      if(linkedForHistory.length)linkedForHistory.forEach(a=>addHistory(a.id,historyText));
      addWorkHistory(w.id,historyText,{status:w.status,statusUpdated:w.updated,kind:statusChanged?'status':'edit'});
      if(prevStatus!==w.status&&isResolvedStatus(w.status))toast('Проблема перемещена во вкладку «Решённые проблемы»');
      const detailWasOpen=!!$('#workDetailOverlay');
      closeEdit();
      if($('#workDetailOverlay'))$('#workDetailOverlay').remove();
      state.selectedWork=detailWasOpen?w.id:null;
      updateNav();queueSave();toast('Работа обновлена и записана в Excel');
      render();
    };
  }
  // Навигация раскрывается только при наведении мыши; состояние не переключается кликом.

  // Стартуем из общей рабочей книги на сервере; sessionStorage не является источником данных.
  (async()=>{
    try{
      const response=await fetch('/api/bootstrap',{cache:'no-store'});const data=await response.json();
      if(!response.ok)throw new Error(data.error||'Не удалось загрузить рабочую книгу');
      loadData(data.sheets,data.name,data.revision,data.appState,data.appRevision);setSyncStatus('Excel синхронизирован','sync-ok');
    }catch(e){
      try{
        loadData(window.DEMO_DATA,'case_data.xlsx','');
        setSyncStatus('Демо-режим · нет связи','sync-warn');
        toast('Не удалось подключиться к рабочей книге: '+e.message,'error');
      }catch(fallbackError){
        showStartupError(fallbackError);
      }
    }
    setInterval(pollWorkbook,2500);
  })();
})();
