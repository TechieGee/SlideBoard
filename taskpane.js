function generateSlides() {
  const text = document.getElementById('lessonInput').value;
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  let slides = [];
  let currentSlide = '', title = '';

  lines.forEach((line) => {
    if (line.match(/^\s*(Grade Level|Subject|Topic|Duration|Lesson Objectives|Prior Knowledge|Introduction|Conclusion|Activity|Assessment|Extension)/i)) {
      if (currentSlide) {
        slides.push({ title, content: currentSlide });
        currentSlide = '';
      }
      title = line;
    } else {
      currentSlide += line + '\n';
    }
  });
  if (currentSlide) {
    slides.push({ title, content: currentSlide });
  }

  slides.forEach(slide => {
    const slideText = slide.title + '\n' + slide.content;
    Office.context.document.setSelectedDataAsync(slideText, { coercionType: 'text' }, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error('Error:', result.error.message);
      }
    });
  });
}
