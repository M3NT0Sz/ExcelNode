const readXlsxFile = require("read-excel-file/node");

const schema = {
  Tipo: {
    prop: "Tipo",
    type: String,
  },
  Enunciado: {
    prop: "Enunciado",
    type: String,
  },
  ImagemEnunciado: {
    prop: "Imagem Enunciado",
    type: String,
  },
  Resposta: {
    prop: "Resposta",
    type: String,
  },
  "Opção A / Imagem A": {
    prop: "Opção A",
    type: String,
  },
  "Opção B / Imagem B": {
    prop: "Opção B",
    type: String,
  },
  "Opção C / Imagem C": {
    prop: "Opção C",
    type: String,
  },
  "Opção D / Imagem D": {
    prop: "Opção D",
    type: String,
  },
};

// File path.
readXlsxFile("./ModeloCreatorAuthor.xlsx", { schema }).then(
  ({ rows, errors }) => {
    console.log(rows);
  }
);

// readXlsxFile("./ModeloCreatorAuthor.xlsx").then((rows) => {
//     console.log(rows)
// });
