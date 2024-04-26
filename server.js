const fs = require('fs')
const Docxtemplater = require('docxtemplater')
const PizZip = require('pizzip')
const XLSX = require('xlsx')

const pathExcel = XLSX.readFile('doc/modelo-excel.xlsx') // Identifica o excel para converter
const nomeColuna = pathExcel.SheetNames[0] // Pega a linha 0 como identificador de nome de colunas

const dataExcel = XLSX.utils.sheet_to_json(pathExcel.Sheets[nomeColuna]) // Converte para JSON

const arrayDeObjetos = dataExcel.map(item => ({ // Transforma em array as colunas
    NOME: item.NOME,
    IDADE: item.IDADE,
    CURSO: item.CURSO
}))

const modelo = './doc/modelo.docx' // Modelo do documento Word

arrayDeObjetos.forEach((data, index) => {
    console.log(`Dados para o documento ${index + 1}:`, data) // Adiciona uma mensagem de log para os dados

    const zip = new PizZip(fs.readFileSync(modelo, 'binary'))
    const docx = new Docxtemplater()
    docx.loadZip(zip)

    docx.setData(data)

    try {
        docx.render()
        console.log(`Documento ${index + 1} renderizado com sucesso!`)
    } catch (error) {
        console.error(`Erro ao renderizar o documento ${index + 1}:`, error)
    }

    let contador = 1
    let documentoAtualizado = `./result/documento_preenchido${index + 1}.docx`

    while (fs.existsSync(documentoAtualizado)) {
        documentoAtualizado = `./result/documento_preenchido${index + 1}_${contador}.docx`
        contador++
    }

    const buffer = docx.getZip().generate({ type: 'nodebuffer' })
    fs.writeFileSync(documentoAtualizado, buffer)
})

console.log('Processo concluído!')