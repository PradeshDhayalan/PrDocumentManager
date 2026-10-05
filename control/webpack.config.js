const path = require('node:path');
// Reuse the platform Fluent exports in the compatibility DatePicker. This also
// preserves the host theme/portal contexts rather than bundling duplicate ones.
const fluent = path.resolve(__dirname, 'DmsGrid/src/services/platformFluent.ts');
module.exports = {resolve:{alias:{
  '@fluentui/react-input$':fluent,
  '@fluentui/react-field$':fluent,
  '@fluentui/react-portal$':fluent,
  '@fluentui/react-tabster$':fluent,
  '@fluentui/react-theme$':fluent,
  '@fluentui/react-shared-contexts$':fluent,
  'react/jsx-runtime$':require.resolve('react/jsx-runtime.js')
}}};
