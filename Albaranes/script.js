document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('albaranForm');
    const linesContainer = document.getElementById('lines');

    document.getElementById('addLine').addEventListener('click', () => {
        const newLine = document.createElement('div');
        newLine.className = 'line';
        newLine.innerHTML = `
            <input type="date" name="fecha" placeholder="Fecha" />
            <input name="cliente" placeholder="Cliente" />
            <input name="itemNumber" placeholder="Item Number" />
            <input name="cantidad" placeholder="Cantidad" />
            <button type="button" class="removeLine">Borrar línea</button>
        `;
        linesContainer.appendChild(newLine);
        addRemoveLineEvent(newLine.querySelector('.removeLine'));
    });

    document.getElementById('resetForm').addEventListener('click', () => {
        form.reset();
        linesContainer.innerHTML = `
            <div class="line">
                <input type="date" name="fecha" placeholder="Fecha" />
                <input name="cliente" placeholder="Cliente" />
                <input name="itemNumber" placeholder="Item Number" />
                <input name="cantidad" placeholder="Cantidad" />
                <button type="button" class="removeLine">Borrar línea</button>
            </div>
        `;
        addRemoveLineEvent(linesContainer.querySelector('.removeLine'));
    });

    document.getElementById('exportExcel').addEventListener('click', () => {
        const lines = Array.from(linesContainer.querySelectorAll('.line')).map(line => {
            return {
                fecha: line.querySelector('input[name="fecha"]').value,
                cliente: line.querySelector('input[name="cliente"]').value,
                itemNumber: line.querySelector('input[name="itemNumber"]').value,
                cantidad: line.querySelector('input[name="cantidad"]').value
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(lines);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Albaranes');
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'albaranes.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    });

    function addRemoveLineEvent(button) {
        button.addEventListener('click', (event) => {
            event.target.parentElement.remove();
        });
    }

    addRemoveLineEvent(linesContainer.querySelector('.removeLine'));
});