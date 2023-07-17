const buttons = document.querySelectorAll(".buttons .btn");
const notifications = document.querySelector(".notifications");

const removeToast = (toast) => {
  toast.classList.add("remove");
  setTimeout(() => toast.remove(), 500);
};

const toastDetails = {
  upsv: {
    icon: "fa-check-circle",
    message: "Đã tải lên danh sách sinh viên!",
  },
  updt: {
    icon: "fa-check-circle",
    message: "Đã tải lên danh sách đề tài",
  },
  cn: {
    icon: "fa-check-circle",
    message: "Cập nhật thành công!",
  },
  xoa: {
    icon: "fa-info-circle",
    message: "Đã xóa thành công !",
  },
  tao: {
    icon: "fa-info-circle",
    message: "Đã tạo trang sheet mới!",
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
