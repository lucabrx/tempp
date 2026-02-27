const TAG_PREFIX = "LINKED_IMAGE_URL:";
const IMAGE_GALLERY = [
  { id: 1, name: "Office View", url: "https://placehold.co/600x400?text=Office" },
  { id: 2, name: "Team Meeting", url: "https://placehold.co/600x400?text=Team" },
  { id: 3, name: "Hello Report", url: "https://placehold.co/600x400?text=Hello" },
];

let selectedImage = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    renderGallery();

    document.getElementById("insertBtn").onclick = insertSelectedImage;
  }
});

function renderGallery() {
  const gallery = document.getElementById("gallery");
  gallery.innerHTML = "";

  IMAGE_GALLERY.forEach((img) => {
    const div = document.createElement("div");
    div.className = "image-item";

    div.onclick = () => {
      document.querySelectorAll(".image-item").forEach((el) => el.classList.remove("selected"));
      div.classList.add("selected");
      selectedImage = img;
      document.getElementById("insertBtn").disabled = false;
    };

    const imageEl = document.createElement("img");
    imageEl.src = img.url;

    const nameEl = document.createElement("div");
    nameEl.className = "image-name";
    nameEl.textContent = img.name;

    div.appendChild(imageEl);
    div.appendChild(nameEl);
    gallery.appendChild(div);
  });
}

async function insertSelectedImage() {
  if (!selectedImage) return;
  const btn = document.getElementById("insertBtn");
  const status = document.getElementById("status");

  btn.disabled = true;
  status.textContent = "Inserting...";

  try {
    const response = await fetch(selectedImage.url);
    const blob = await response.blob();
    const base64WithPrefix = await blobToBase64(blob);
    const base64Data = base64WithPrefix.split(",")[1];

    await Word.run(async (context) => {
      const range = context.document.getSelection();
      const cc = range.insertContentControl();
      cc.tag = TAG_PREFIX + selectedImage.url;
      cc.insertInlinePictureFromBase64(base64Data, Word.InsertLocation.replace);
      await context.sync();
    });

    status.textContent = "✓ Success!";
    status.style.color = "green";
  } catch (error) {
    status.textContent = "Error: " + error.message;
    status.style.color = "red";
  } finally {
    btn.disabled = false;
  }
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
