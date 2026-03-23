// ===== ACTIVE MENU =====
document.addEventListener("DOMContentLoaded", function () {
    const links = document.querySelectorAll(".sidebar a");
    const currentPath = window.location.pathname;

    links.forEach(link => {
        if (link.getAttribute("href") === currentPath) {
            link.parentElement.classList.add("active");
        }
    });
});

// ===== XÁC NHẬN XOÁ =====
function xacNhanXoa() {
    return confirm("Bạn có chắc chắn muốn xoá không?");
}

// ===== HIỆN / ẨN FORM =====
function toggleForm(id) {
    const form = document.getElementById(id);
    if (form.style.display === "none" || form.style.display === "") {
        form.style.display = "block";
    } else {
        form.style.display = "none";
    }
}
