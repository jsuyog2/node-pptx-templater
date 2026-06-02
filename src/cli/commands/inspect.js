/**
 * @fileoverview `inspect` CLI command — detailed PPTX structure inspection.
 */
const chalk = require('chalk')
const ora = require('ora')
const { resolve } = require('path')
const { PPTXTemplater } = require('../../index.js')

async function inspectCommand(filePath, opts) {
  const showAll = opts.all
  const spinner = ora(`Inspecting: ${filePath}`).start()

  try {
    const ppt = await PPTXTemplater.load(resolve(filePath))
    const info = ppt.getInfo()
    spinner.stop()

    console.log(chalk.bold.cyan('\n═══ PPTX Inspection Report ═══\n'))
    console.log(chalk.bold('General:'))
    console.log(`  Title:    ${info.title || chalk.dim('(none)')}`)
    console.log(`  Author:   ${info.author || chalk.dim('(none)')}`)
    console.log(`  Created:  ${info.created || chalk.dim('(unknown)')}`)
    console.log(`  Slides:   ${chalk.cyan(info.slideCount)}`)
    console.log(`  Media:    ${chalk.cyan(info.mediaCount)} files`)

    if (opts.slides || showAll) {
      console.log(chalk.bold('\nSlides:'))
      for (const slide of info.slides) {
        const tags = slide.tags.length > 0 ? chalk.dim(` [${slide.tags.join(', ')}]`) : ''
        console.log(`  ${chalk.cyan(slide.index.toString().padStart(2))}. ${slide.zipPath}${tags}`)
      }
    }

    console.log('')
  } catch (err) {
    spinner.fail(chalk.red(`Inspect failed: ${err.message}`))
    process.exit(1)
  }
}

module.exports = { inspectCommand }
