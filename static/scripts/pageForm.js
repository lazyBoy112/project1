const buttons = document.querySelectorAll(".buttons .btn");
const notifications = document.querySelector(".notifications");

const removeToast = (toast) => {
  toast.classList.add("remove");
  setTimeout(() => toast.remove(), 500);
};

const toastDetails = {
    taoform: {
        icon: "fa-check-circle",
        message: "Đã tạo được form !",
      },
      copy: {
        icon: "fa-check-circle",
        message: "Đã copy link form!",
      },
      tai: {
        icon: "fa-check-circle",
        message: "Đã tải form thành công !",
      },
    
};
const handleCreateToast = (id) => {
  const { icon, message } = toastDetails[id];

  const toast = document.createElement("li");
  toast.className = `toast ${id}`;
  toast.innerHTML = `
  <div class="column">
          <i class="fa ${icon}"></i>
          <span>${message}</span>
        </div>
        <i class="fa-solid fa-xmark" onclick="removeToast(this.parentElement)"></i>
  `;
  notifications.appendChild(toast);
  setTimeout(() => removeToast(toast), 5000);
};
buttons.forEach((button) => {
  button.addEventListener("click", () => {
    handleCreateToast(button.id);
  });
});
