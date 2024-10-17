const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

// Função para encontrar o índice de uma coluna específica
function findColumnIndex(headers, columnName) {
    return headers.findIndex(header => header.toLowerCase().includes(columnName.toLowerCase()));
}

// Função para validar email
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Mapeamento de símbolos para letras correspondentes
const symbolMap = {
    '@': 'a',
    '$': 's',
    '#': 'h',
    '1': 'l',
    '&': 'e',
    '3': 'e',
    '4': 'a',
    '5': 's',
    '7': 't',
    '!': 'i',
    '%': 'o',
    '√': 'v',
    'º': 'o'
};

// Função para substituir símbolos por letras semelhantes
function replaceSymbolsWithLetters(name) {
    return name
        .split('') // Dividir o nome em caracteres individuais
        .map((char, index) => {
            if (symbolMap[char]) {
                // Se for o primeiro caractere, retorna a letra em maiúsculo
                return index === 0 ? symbolMap[char].toUpperCase() : symbolMap[char].toLowerCase();
            }
            return char; // Se não for um símbolo mapeado, mantém o caractere original
        })
        .join(''); // Junta os caracteres de volta em uma string
}

// Função para normalizar nomes (remover acentos, substituir símbolos)
function normalizeName(name) {
    return name
        .normalize('NFD') // Decompor caracteres acentuados
        .replace(/[\u0300-\u036f]/g, '') // Remover acentos
        .replace(/[^\w\s\-\']/g, '') // Remover qualquer símbolo não desejado
        .trim(); // Remover espaços no início e no final
}

// Função para processar o nome completo
function processFullName(firstName, middleName, lastName) {
    let fullName = [firstName, middleName, lastName].filter(Boolean).join(' ').trim(); // Concatena sem espaços extras
    fullName = replaceSymbolsWithLetters(fullName); // Substituir símbolos por letras correspondentes
    fullName = normalizeName(fullName); // Normalizar o nome (remover acentos e símbolos indesejados)
    return fullName;
}

// Função para normalizar o nome
function normalizaName(name) {
    return name
        .split(' ') // Dividir o nome em palavras
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()) // Primeira letra maiúscula, resto minúsculo
        .join(' '); // Juntar as palavras de volta com espaços
}

// Função para normalizar o número de telefone no formato (XX) XXXXX-XXXX
function formatPhoneNumber(phone) {
    if (!phone) return '';

    // Limpar o telefone, removendo caracteres que não sejam números
    let cleanedPhone = phone.replace(/[^\d]/g, '');

    // Caso o número tenha prefixo +55, removê-lo
    if (cleanedPhone.startsWith('55')) {
        cleanedPhone = cleanedPhone.slice(2);
    }
    
    // Caso o número tenha prefixo 03135, removê-lo
    if (cleanedPhone.startsWith('03135')) {
        cleanedPhone = cleanedPhone.slice(5);  // Remove o "03135"
    }

    // Se o número tiver 11 dígitos, com DDD e nono dígito
    if (cleanedPhone.length === 11) {
        return `(${cleanedPhone.slice(0, 2)}) ${cleanedPhone.slice(2, 7)}-${cleanedPhone.slice(7)}`;
    }
    // Se o número tiver 10 dígitos, com DDD e sem nono dígito
    else if (cleanedPhone.length === 10) {
        return `(${cleanedPhone.slice(0, 2)}) ${cleanedPhone.slice(2, 6)}-${cleanedPhone.slice(6)}`;
    }
    // Se o número tiver 9 dígitos, sem DDD e com nono dígito
    else if (cleanedPhone.length === 9) {
        return `${cleanedPhone.slice(0, 5)}-${cleanedPhone.slice(5)}`;
    }
    // Se o número tiver 8 dígitos, sem DDD e sem nono dígito
    else if (cleanedPhone.length === 8) {
        return `${cleanedPhone.slice(0, 4)}-${cleanedPhone.slice(4)}`;
    }

    // Se não for nenhum dos formatos, retornar vazio
    return '';
}

// Definir estilo de fonte Arial e tamanho 10
const cellStyle = {
    font: { name: 'Arial', sz: 10 },
};

// Função para ajustar a largura das colunas com base no conteúdo
function adjustColumnWidths(worksheet) {
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    const colWidths = [];

    for (let col = range.s.c; col <= range.e.c; col++) {
        let maxLength = 0;

        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = xlsx.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];

            if (cell && cell.v) {
                maxLength = Math.max(maxLength, String(cell.v).length);
            }
        }

        colWidths[col] = { width: maxLength + 2 }; // Adicionar um pequeno espaço para garantir que o texto não corte
    }

    worksheet['!cols'] = Object.keys(colWidths).map(colIndex => colWidths[colIndex]);
}

// Rota para upload e conversão do arquivo CSV
app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    // Caminho do arquivo CSV carregado
    const csvFilePath = req.file.path;

    // Ler o arquivo CSV
    fs.readFile(csvFilePath, 'utf8', (err, data) => {
        if (err) {
            return res.status(500).send('Error reading the CSV file.');
        }

        // Processar o arquivo CSV
        const lines = data.split('\n');
        const headers = lines[0].split(',').map(header => header.trim());

        // Encontrar os índices das colunas de interesse
        const firstNameIndex = findColumnIndex(headers, 'First Name');
        const middleNameIndex = findColumnIndex(headers, 'Middle Name');
        const lastNameIndex = findColumnIndex(headers, 'Last Name');
        const phone1Index = findColumnIndex(headers, 'Phone 1 - Value');
        const phone2Index = findColumnIndex(headers, 'Phone 2 - Value');
        const emailIndex = findColumnIndex(headers, 'E-mail 1 - Value');

        const contacts = [];

        // Percorrer todas as linhas (exceto o cabeçalho) e extrair as informações
        for (let i = 1; i < lines.length; i++) {
            const columns = lines[i].split(',').map(column => column.trim());

            // Montar o nome completo e processar símbolos
            const firstName = columns[firstNameIndex] || '';
            const middleName = columns[middleNameIndex] || '';
            const lastName = columns[lastNameIndex] || '';

            let fullName = normalizaName(processFullName(firstName, middleName, lastName));

            // Verificar se o nome completo está vazio e tentar preenchê-lo com telefone ou e-mail
            if (!fullName) {
                const phone1 = columns[phone1Index] || '';
                const email = columns[emailIndex] || '';
                
                if (phone1) {
                    fullName = formatPhoneNumber(phone1);  // Se houver telefone, colocar no campo do nome com formatação
                } else if (email) {
                    fullName = email;  // Se não houver telefone, colocar e-mail no campo do nome
                }
            }

            // Limpar e formatar os telefones
            let phone1 = formatPhoneNumber(columns[phone1Index] || '');
            let phone2 = '';

            // Verificar se há separador ":::"
            if (phone1.includes(':::')) {
                const phoneParts = phone1.split(':::');
                phone1 = formatPhoneNumber(phoneParts[0]); // Primeiro telefone
                phone2 = formatPhoneNumber(phoneParts[1] || ''); // Segundo telefone, se existir
            } else {
                // Se não houver separador, apenas formatar o segundo telefone normalmente
                phone2 = formatPhoneNumber(columns[phone2Index] || '');
            }

            // Verificar se o email é válido
            const email = isValidEmail(columns[emailIndex]) ? columns[emailIndex] : '';

            // Adicionar contato ao array, mesmo que esteja faltando o email ou telefone
            if (fullName) {
                contacts.push({ fullName, phone1, phone2, email });
            }
        }

        // Criar um novo arquivo XLSX com os dados processados
        const workbook = xlsx.utils.book_new();
        const worksheetData = [['Nome', 'Telefone 1', 'Telefone 2', 'E-mail'], ...contacts.map(c => [c.fullName, c.phone1, c.phone2, c.email])];
        const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);

        // Ajustar a largura das colunas
        adjustColumnWidths(worksheet);

        // Aplica o estilo de fonte Arial e tamanho 10 em todas as células
        const range = xlsx.utils.decode_range(worksheet['!ref']);
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = xlsx.utils.encode_cell({ r: row, c: col });
                if (!worksheet[cellAddress]) continue;
                worksheet[cellAddress].s = cellStyle;
            }
        }

        xlsx.utils.book_append_sheet(workbook, worksheet, 'Contacts');

        const outputPath = path.join(__dirname, 'uploads', 'import-contatos.xlsx');
        xlsx.writeFile(workbook, outputPath);

        // Enviar o arquivo XLSX gerado de volta para o cliente
        res.download(outputPath, 'import-contatos.xlsx', (err) => {
            if (err) {
                console.error('Error downloading the file:', err);
            }
            // Remover o arquivo CSV e o XLSX gerado após o download
            fs.unlink(csvFilePath, () => {});
            fs.unlink(outputPath, () => {});
        });
    });
});

// Iniciar o servidor
app.listen(3000, () => {
    console.log('Server running on http://localhost:3000');
});
