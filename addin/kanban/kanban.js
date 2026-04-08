let currentTasks = [];

Office.onReady(init);

async function init() {
  currentTasks = await loadTasks();
  render();
  setupDnD();
}

function render() {
  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");
  currentTasks.sort((a,b)=>a.order-b.order).forEach(task=>{
    const el=document.createElement("div");
    el.className="card";
    el.draggable=true;
    el.textContent=task.name;
    el.dataset.id=task.id;
    el.dataset.row=task.row;
    el.addEventListener("dragstart",e=>{
      el.classList.add("dragging");
      e.dataTransfer.setData("id",task.id);
    });
    el.addEventListener("dragend",()=>el.classList.remove("dragging"));
    document.querySelector(`#${task.status} .card-list`).appendChild(el);
  });
}

function setupDnD(){
  document.querySelectorAll(".lane").forEach(lane=>{
    lane.addEventListener("dragover",e=>{
      e.preventDefault();
      const container=lane.querySelector(".card-list");
      const after=getAfter(container,e.clientY);
      const dragging=document.querySelector(".dragging");
      if(!after) container.appendChild(dragging);
      else container.insertBefore(dragging,after);
    });
    lane.addEventListener("drop",async e=>{
      const id=e.dataTransfer.getData("id");
      const task=currentTasks.find(t=>t.id==id);
      task.status=lane.id;
      await saveOrder(lane.id);
      render();
    });
  });
}

function getAfter(container,y){
  const els=[...container.querySelectorAll(".card:not(.dragging)")];
  return els.reduce((closest,child)=>{
    const box=child.getBoundingClientRect();
    const offset=y-box.top-box.height/2;
    if(offset<0 && offset>closest.offset){
      return {offset,element:child};
    }else return closest;
  },{offset:Number.NEGATIVE_INFINITY}).element;
}
