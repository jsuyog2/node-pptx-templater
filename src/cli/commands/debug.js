/**
 * @fileoverview `debug` CLI command — diagnoses and repairs corrupted PPTX files.
 */
const chalk = require('chalk')
const ora = require('ora')
const { resolve } = require('path')
const { PPTXTemplater } = require('../../index.js')

async function debugCommand(filePath, opts) {
  const spinner = ora(`Loading PPTX for debug: ${filePath}`).start()

  try {
    const ppt = await PPTXTemplater.load(resolve(filePath))
    spinner.succeed('Loaded successfully')

    const result = ppt.validate()

    console.log(chalk.bold.cyan('\n═══ Debug Report ═══\n'))
    console.log(chalk.bold('Validation:'))

    if (result.valid) {
      console.log(chalk.green('  ✓ Structure is valid'))
    } else {
      console.log(chalk.red(`  ✗ ${result.errors.length} error(s) found`))
      result.errors.forEach(e => console.log(chalk.red(`    • ${e}`)))
    }

    if (result.warnings.length > 0) {
      console.log(chalk.yellow(`\n  ${result.warnings.length} warning(s):`))
      result.warnings.forEach(w => console.log(chalk.yellow(`    • ${w}`)))
    }

    if (opts.fix && opts.out) {
      spinner.start('Attempting repairs...')
      await ppt.saveToFile(resolve(opts.out))
      spinner.succeed(chalk.green(`✓ Repaired PPTX saved to: ${opts.out}`))
    }

    console.log('')
  } catch (err) {
    spinner.fail(chalk.red(`Debug failed: ${err.message}`))
    if (process.env.PPTX_LOG_LEVEL === 'debug') console.error(err.stack)
    process.exit(1)
  }
}

module.exports = { debugCommand }
