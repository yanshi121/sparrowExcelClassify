function Heartbeat(url) {
    setInterval(()=>{
        return fetch(url)
        .then(response => {
            if (!response.ok) {
                console.log('Network response was not ok');
            }
            return response.json();
        })
        .catch(error => {
            alert('There has been a problem with your fetch operation:');
        });
    })
}