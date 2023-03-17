var doc = new jspdf.jsPDF("a4");

const downloadButton = document.getElementById('btn-download');
      downloadButton.addEventListener('click', function() {
        var image = window.print();
        doc.save(image,'tabela.pdf');
      });