const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

function encontrarIndiceColuna(cabecalhos, nomeColuna) {
    return cabecalhos.findIndex(cabecalho => cabecalho.toLowerCase().includes(nomeColuna.toLowerCase()));
} //verifica se o cabeçalho que eu estou verificando bate com o nome da coluna que eu quero 
//a busca é case-insensitive (não diferencia maiúsculas e minúsculas)

function emailValido(email) {
    const regexEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return regexEmail.test(email);
} //Verifica se o email é válido usando uma expressão regular.

const mapaSimbolos = {
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

function substituirSimbolosPorLetras(nome) {
    return nome
        .split('') //A função split('') divide a string nome em um array de caracteres individuais.
        .map((char, index) => {
            if (mapaSimbolos[char]) {
                return index === 0 ? mapaSimbolos[char].toUpperCase() : mapaSimbolos[char].toLowerCase();
            }//ele transforma o char no char do mapa (se tiver) e então é esse char do mapa que é colocado como minúscula ou maiúscula
            return char;
        })
        .join('');
}

function normalizarNome(nome) {// Remove caracteres especiais indesejados do nome
    return nome
        .replace(/[^\w\s\-çãáéíóúâêîôûãõäëïöü]/g, '') // Permitir 'ç' e outros caracteres acentuados
        .trim();// o trim(); exclui espaços desnecessários  
}

function processarNomeCompleto(primeiroNome, nomeDoMeio, sobrenome) {
    let nomeCompleto = [primeiroNome, nomeDoMeio, sobrenome].filter(Boolean).join(' ').trim();
    nomeCompleto = substituirSimbolosPorLetras(nomeCompleto);
    nomeCompleto = normalizarNome(nomeCompleto);
    return nomeCompleto;
}/*filter(Boolean): Remove qualquer valor falso (como null, undefined, ou strings vazias) do array. 
 Assim, se qualquer um dos 3 for vazio ou undefined, ele será removido.*/

function normalizarNomeCompleto(nome) {
    return nome
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
        .join(' ');
}// o método split(' ') separa a string nome em palavras individuais, criando um array de palavras.
/*Não é necessário declarar explicitamente word como uma variável antes porque estamos usando uma arrow 
function.
O map() passa cada elemento do array para a função de callback, e o parâmetro (word) é o nome que 
escolhemos para representar cada elemento.
Com a arrow function (=>) você pode omitir a palavra-chave function.
*/

function formatarTelefone(telefone) {
    if (!telefone) return '';

    let telefoneLimpo = telefone.replace(/[^\d]/g, '');

    if (telefoneLimpo.startsWith('55')) {
        telefoneLimpo = telefoneLimpo.slice(2);
    }

    if (telefoneLimpo.startsWith('031')) {
        telefoneLimpo = telefoneLimpo.slice(3);
    }

    if (telefoneLimpo.length === 11) {
        return `(${telefoneLimpo.slice(0, 2)}) ${telefoneLimpo.slice(2, 7)}-${telefoneLimpo.slice(7)}`;
    }
    else if (telefoneLimpo.length === 10) {
        return `(${telefoneLimpo.slice(0, 2)}) ${telefoneLimpo.slice(2, 6)}-${telefoneLimpo.slice(6)}`;
    }
    else if (telefoneLimpo.length === 9) {
        return `${telefoneLimpo.slice(0, 5)}-${telefoneLimpo.slice(5)}`;
    }
    else if (telefoneLimpo.length === 8) {
        return `${telefoneLimpo.slice(0, 4)}-${telefoneLimpo.slice(4)}`;
    }

    return telefone; // Se o telefone não puder ser formatado, retorna o valor original
}

function processarTelefones(telefone1Raw, telefone2Raw) {
    // Se já houver um número na coluna Phone 2 - Value, usá-lo como telefone2
    if (telefone2Raw) {
        const telefone2 = formatarTelefone(telefone2Raw);
        const telefone1 = formatarTelefone(telefone1Raw);
        return { telefone1, telefone2 };
    }

    // Se não houver número em Phone 2 - Value, verifica se Phone 1 - Value contém ":::"
    if (telefone1Raw.includes(':::')) {
        const telefones = telefone1Raw.split(':::').map(tel => tel.trim());

        const telefone1 = formatarTelefone(telefones[0]); // Primeiro telefone antes do ":::"
        const telefone2 = telefones[1] ? formatarTelefone(telefones[1]) : ''; // Segundo telefone após o ":::", se existir

        return { telefone1, telefone2 };
    }

    // Se não houver ":::" e nem Phone 2 - Value, apenas formata Phone 1
    const telefone1 = formatarTelefone(telefone1Raw);
    return { telefone1, telefone2: '' }; // Retorna telefone2 vazio
}


app.post('/converter', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('Nenhum arquivo enviado.');
    }

    const caminhoCSV = req.file.path;

    //linha 130 - 135:
    /*app.post('/converter', ...): Define a rota que responde a requisições POST no caminho /converter.
      upload.single('file'): Middleware que trata o upload de um único arquivo enviado no campo 'file' do 
      formulário. Ele salva temporariamente o arquivo e o disponibiliza em req.file.
      (req, res) => {}: Função de callback que é executada quando a rota é chamada. Ela verifica se o arquivo 
      foi enviado e inicia o processamento.*/

    fs.readFile(caminhoCSV, 'utf8', (err, data) => {
        if (err) {
            return res.status(500).send('Erro ao ler o arquivo CSV.');
        }

        const linhas = data.split('\n');
        const cabecalhos = linhas[0].split(',').map(cabecalho => cabecalho.trim());

        //linha 144 - 150: 
        /*O arquivo CSV é lido como uma string (com todas as linhas juntas).
        Essa string é dividida em linhas usando split('\n').
        A primeira linha (linhas[0]) é extraída e dividida em colunas usando split(',').
        Cada coluna (ou seja, cada cabeçalho) é "limpa" de espaços extras usando map(cabecalho => cabecalho.trim()).
        O resultado final é o array cabecalhos, que contém os nomes das colunas como strings.*/

        const indicePrimeiroNome = encontrarIndiceColuna(cabecalhos, 'First Name');
        const indiceNomeDoMeio = encontrarIndiceColuna(cabecalhos, 'Middle Name');
        const indiceSobrenome = encontrarIndiceColuna(cabecalhos, 'Last Name');
        const indiceTelefone1 = encontrarIndiceColuna(cabecalhos, 'Phone 1 - Value');
        const indiceTelefone2 = encontrarIndiceColuna(cabecalhos, 'Phone 2 - Value');
        const indiceEmail = encontrarIndiceColuna(cabecalhos, 'E-mail 1 - Value');
        const indiceOrganizacao = encontrarIndiceColuna(cabecalhos, 'Organization Name'); // Novo índice

        const contatos = [];

        for (let i = 1; i < linhas.length; i++) {
            const colunas = linhas[i].split(',').map(coluna => coluna.trim());

            const primeiroNome = colunas[indicePrimeiroNome] || '';
            const nomeDoMeio = colunas[indiceNomeDoMeio] || '';
            const sobrenome = colunas[indiceSobrenome] || '';
            const organizacao = colunas[indiceOrganizacao] || '';

            let nomeCompleto = normalizarNomeCompleto(processarNomeCompleto(primeiroNome, nomeDoMeio, sobrenome));

            if (!nomeCompleto) {
                const telefone1 = colunas[indiceTelefone1] || '';
                const email = colunas[indiceEmail] || '';

                if (telefone1) {
                    nomeCompleto = formatarTelefone(telefone1);
                } else if (email) {
                    nomeCompleto = email;
                }
            }

            // Usa a nova função para processar telefone 1 e telefone 2, dando prioridade para Phone 2 - Value
            const { telefone1, telefone2 } = processarTelefones(colunas[indiceTelefone1] || '', colunas[indiceTelefone2] || '');

            const email = colunas[indiceEmail] && emailValido(colunas[indiceEmail]) ? colunas[indiceEmail] : '';
            const organizacaoFormatada = organizacao ? `${nomeCompleto} | ${organizacao}` : '';

            if (nomeCompleto || telefone1 || email || organizacao) {
                contatos.push([nomeCompleto, telefone1, telefone2, email, organizacaoFormatada]); // Adicionando a organização como nova coluna
            }
        }

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Contatos');

        // Definir estilo de célula com fonte Arial e tamanho 10
        const estiloCelula = {
            font: { name: 'Arial', size: 10 }
        };

        // Adicionar cabeçalhos
        sheet.addRow(['Nome', 'Telefone 1', 'Telefone 2', 'E-mail', 'Organização']).eachCell(cell => { // Adicionar cabeçalho para a nova coluna
            cell.style = estiloCelula;
        });

        // Adicionar dados dos contatos
        contatos.forEach(contato => {
            sheet.addRow(contato).eachCell(cell => {
                cell.style = estiloCelula;
            });
        });

        // Ajustar largura das colunas
        sheet.columns.forEach(column => {
            const maxLength = column.values.reduce((max, val) => {
                return Math.max(max, (val ? String(val).length : 0));
            }, 0);
            column.width = maxLength + 1; // Adiciona um espaço extra
        });

        const caminhoExcel = path.join(__dirname, 'uploads', 'contatos_formatados.xlsx');

        workbook.xlsx.writeFile(caminhoExcel)
            .then(() => {
                fs.unlinkSync(caminhoCSV);

                res.download(caminhoExcel, 'contatos_formatados.xlsx', (err) => {
                    if (err) {
                        return res.status(500).send('Erro ao enviar o arquivo Excel.');
                    }

                    fs.unlinkSync(caminhoExcel);
                });
            })
            .catch(err => {
                console.error(err);
                res.status(500).send('Erro ao gerar o arquivo Excel.');
            });
    });
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
    console.log(`Servidor iniciado na porta ${port}`);
});