
document.getElementById('templateBtn1').addEventListener('click', function() {
// Toggle 'active' class on button
this.classList.toggle('active');
});



document.getElementById('templateBtn2').addEventListener('click', function() {
// Toggle 'active' class on button
this.classList.toggle('active');
});



document.getElementById('templateBtn3').addEventListener('click', function() {
// Toggle 'active' class on button
this.classList.toggle('active');
});


let time = document.getElementById("current-time");
let sunSvg = document.getElementById("sun-svg");
let moonSvg = document.getElementById("moon-svg");

setInterval(() => {
    let d = new Date();
    let hours = d.getHours();
    if (hours >= 7 && hours < 19) {
        sunSvg.style.display = "block";
        moonSvg.style.display = "none";
    } else {
        sunSvg.style.display = "none";
        moonSvg.style.display = "block";
    }

    time.innerHTML = d.toLocaleTimeString();
}, 1000);




const dropdown = document.getElementById("dropdown");
const dropdownContent = document.getElementById("dropdownContent");

function toggleDropdown() {
    dropdownContent.style.display = dropdownContent.style.display === "block" ? "none" : "block";
}

function handleCheckboxChange(checkbox) {
    const selectedItemsDiv = document.getElementById("selectedItems");
    const label = checkbox.nextElementSibling.textContent;

    if (checkbox.checked) {
        const selectedItemDiv = document.createElement("div");
        selectedItemDiv.className = "selected-item";
        selectedItemDiv.textContent = label;

        const cutBtn = document.createElement("span");
        cutBtn.className = "cut-btn";
        cutBtn.textContent = "X";
        cutBtn.onclick = function () {
            selectedItemDiv.remove();
            checkbox.checked = false;
        };

        selectedItemDiv.appendChild(cutBtn);
        selectedItemsDiv.appendChild(selectedItemDiv);
    } else {
        const selectedItems = selectedItemsDiv.getElementsByClassName("selected-item");
        for (const selectedItem of selectedItems) {
            if (selectedItem.textContent === label) {
                selectedItem.remove();
                break;
            }
        }
    }
}


document.addEventListener("click", function(event) {
    if (!dropdown.contains(event.target)) {
        dropdownContent.style.display = "none";
    }
});


document.getElementById('filterhistory').addEventListener('click', function() {
    this.classList.toggle('active');
    });


    