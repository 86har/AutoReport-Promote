
setTimeout(()=>{
    Array.from(document.getElementsByClassName('report-item')).forEach((reportItem, i)=>{
        console.log(reportItem)
        reportItem.addEventListener('click', ()=>{
            console.log(`clicked ${i}`)
        })
    })
}, 400)
