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
    message: "đã cập nhật thành công !",
  },
  // xoa: {
  //   icon: "fa-info-circle",
  //   message: "Đã xóa thành công !",
  // },
  tao: {
    icon: "fa-info-circle",
    message: "Đã tạo trang danh sách đăng ký đề tài!",
  },
};
const handleCreateToast = (id) => {
  const { icon, message } = toastDetails[id];

  const loadingToast = document.createElement("li");
  loadingToast.className = "toast loading";
  loadingToast.innerHTML = `
    <div class="column">
      <i class="fa fa-spinner fa-spin"></i>
      <span>Loading...</span>
    </div>
  `;
  notifications.appendChild(loadingToast);

  setTimeout(() => {
    removeToast(loadingToast);

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
  }, 5000);
};

buttons.forEach((button) => {
  button.addEventListener("click", () => {
    handleCreateToast(button.id);
  });
});
function displaySelectedFile(inputId, displayId) {
  const fileInput = document.getElementById(inputId);
  const fileNameDisplay = document.getElementById(displayId);

  if (fileInput.files.length > 0) {
    const fileName = fileInput.files[0].name;
    fileNameDisplay.textContent = fileName;
  } else {
    fileNameDisplay.textContent = "Chưa chọn tệp";
  }
}

