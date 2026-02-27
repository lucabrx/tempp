const TAG_PREFIX = "LINKED_IMAGE_URL:";
const IMAGE_GALLERY = [
  { id: 1, name: "Office View", url: "https://placehold.co/600x400" },
  { id: 2, name: "Team Meeting", url: "https://placehold.co/600x400" },
  { id: 3, name: "Graph Report", url: "https://placehold.co/600x400" },
];

let selectedImage = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    renderGallery();

    document.getElementById("insertBtn").onclick = insertSelectedImage;
    document.getElementById("refreshBtn").onclick = refreshAllImages;
  }
});

async function insertSelectedImage() {
  if (!selectedImage) return;

  const btn = document.getElementById("insertBtn");
  const status = document.getElementById("status");

  btn.disabled = true;
  status.textContent = "Processing image...";

  try {
    const response = await fetch(selectedImage.url);
    if (!response.ok) throw new Error("Network response was not ok");
    const blob = await response.blob();
    const base64WithPrefix = await blobToBase64(blob);
    const base64Data = base64WithPrefix.split(",")[1];

    await Word.run(async (context) => {
      const range = context.document.getSelection();

      const cc = range.insertContentControl();
      cc.tag = TAG_PREFIX + selectedImage.url;
      cc.title = "Linked: " + selectedImage.name;
      cc.appearance = "BoundingBox";

      cc.insertInlinePictureFromBase64(base64Data, Word.InsertLocation.replace);

      await context.sync();
    });

    status.textContent = "✓ Inserted!";
    status.style.color = "green";
  } catch (error) {
    status.textContent = "Error: " + error.message;
    status.style.color = "red";
  } finally {
    btn.disabled = false;
  }
}

async function refreshAllImages() {
  const status = document.getElementById("status");
  status.textContent = "Syncing with server...";

  try {
    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag");
      await context.sync();

      const targets = controls.items.filter((cc) => (cc.tag || "").startsWith(TAG_PREFIX));

      for (const cc of targets) {
        const url = cc.tag.substring(TAG_PREFIX.length);
        const response = await fetch(url);
        const blob = await response.blob();
        const base64WithPrefix = await blobToBase64(blob);
        const base64Data = base64WithPrefix.split(",")[1];

        cc.insertInlinePictureFromBase64(base64Data, Word.InsertLocation.replace);
      }
      await context.sync();
    });
    status.textContent = "✓ All images updated.";
  } catch (e) {
    status.textContent = "Refresh failed: " + e.message;
  }
}

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

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
