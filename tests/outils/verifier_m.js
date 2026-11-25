// Vérifie la syntaxe des requêtes Power Query (M) avec l'analyseur officiel de Microsoft
// (@microsoft/powerquery-parser). Usage : node verifier_m.js <chemin_parser> fichier1.pq ...
const parserPath = process.argv[2];
const PQP = require(parserPath);
const fs = require("fs");
(async () => {
  let erreurs = 0;
  for (const f of process.argv.slice(3)) {
    const texte = fs.readFileSync(f, "utf8").replace(/^﻿/, "");
    const task = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, texte);
    if (PQP.TaskUtils.isOk(task)) {
      console.log("OK     " + f);
    } else {
      erreurs++;
      const err = task.error;
      console.log("ERREUR " + f + " : " + (err && err.message ? err.message : JSON.stringify(err)).slice(0, 600));
    }
  }
  process.exit(erreurs ? 1 : 0);
})();
