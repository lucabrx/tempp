const IMAGE_GALLERY = [
  {
    id: 1,
    name: "Image 1",
    url: "https://placehold.co/600x400",
    base64: null 
  },
  {
    id: 2,
    name: "Image 2",
    url: "https://placehold.co/600x400",
    base64: null,
  }
];

let selectedImage = null;

Office.onReady(() => {
  renderGallery();
  
  document.getElementById('insertBtn').onclick = async () => {
    if (!selectedImage) return;
    
    const btn = document.getElementById('insertBtn');
    const status = document.getElementById('status');
    btn.disabled = true;
    status.textContent = "Inserting...";
    
    try {
      let base64Data;
      
      if (selectedImage.base64) {
        base64Data = selectedImage.base64;
      } else {
        status.textContent = "Loading image...";
        const response = await fetch(selectedImage.url);
        if (!response.ok) throw new Error('Failed to load image');
        const blob = await response.blob();
        const dataUrl = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(reader.result);
          reader.onerror = reject;
          reader.readAsDataURL(blob);
        });
        base64Data = dataUrl.split(',')[1];
      }
      
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertInlinePictureFromBase64(base64Data, Word.InsertLocation.replace);
        await context.sync();
      });
      
      status.textContent = "✓ Inserted successfully!";
      status.style.color = "green";
      
    } catch (error) {
      console.error(error);
      status.textContent = "Error: " + error.message;
      status.style.color = "red";
      btn.disabled = false;
    }
  };
});

function renderGallery() {
  const gallery = document.getElementById('gallery');
  gallery.innerHTML = '';
  
  IMAGE_GALLERY.forEach(img => {
    const div = document.createElement('div');
    div.className = 'image-item';
    div.onclick = () => selectImage(img.id, div);
    
    const imageEl = document.createElement('img');
    imageEl.src = img.url;
    
    const nameEl = document.createElement('div');
    nameEl.className = 'image-name';
    nameEl.textContent = img.name;
    
    div.appendChild(imageEl);
    div.appendChild(nameEl);
    gallery.appendChild(div);
  });
}

function selectImage(id, element) {
  document.querySelectorAll('.image-item').forEach(el => {
    el.classList.remove('selected');
  });
  
  element.classList.add('selected');
  selectedImage = IMAGE_GALLERY.find(img => img.id === id);
  
  document.getElementById('insertBtn').disabled = false;
  document.getElementById('status').textContent = "";
}