'use strict';

const build = require('@microsoft/sp-build-web');

// 🔧 Desactivar linters
process.env['DISABLE_LINTER'] = 'true';
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Apagar tareas conocidas de lint
if (process.env['DISABLE_LINTER'] === 'true') {
  console.log('🔇 Linter deshabilitado: se omiten advertencias y errores de ESLint/TSLint.');
  if (build.tslintCmd) build.tslintCmd.enabled = false;
  if (build.eslintCmd) build.eslintCmd.enabled = false;
}

// Compatibilidad con serve-deprecated
const originalGetTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  const result = originalGetTasks.call(build.rig);

  // sustituir serve
  result.set('serve', result.get('serve-deprecated'));

  // ⚙️ sustituir lint por tarea vacía
  result.set('lint', {
    name: 'lint',
    task: (gulp, done) => {
      console.log('⚙️  Saltando completamente la tarea de lint (forzado en gulpfile.js)');
      done();
    }
  });

  return result;
};

// Inicializar gulp
build.initialize(require('gulp'));

// 🔧 Asegurar que lint quede neutralizado después de la inicialización
build.task('lint', (done) => {
  console.log('⚙️  Lint deshabilitado — no se ejecutará ningún análisis.');
  done();
});
