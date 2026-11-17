/**
 * @fileoverview `extract` CLI command — extracts XML parts, folder trees, or ZIP packages from a PPTX.
 */
const chalk = require('chalk')
const { resolve } = require('path')
const { writeFile } = require('fs/promises')
const { PPTXTemplater } = require('../../index.js')
const { formatXml } = require('../../writer/OutputWriter.js')

async function extractCommand(filePath, opts) {
  try {
    const ppt = await PPTXTemplater.load(resolve(filePath))

    // 1. Extract entire presentation to uncompressed XML directory
    if (opts.all || opts.folder) {
      const targetDir = resolve(opts.folder || opts.out || './extracted')
      const format = Boolean(opts.format || opts.formatXml)
      await ppt.extractToXml(targetDir, { formatXml: format })
      console.log(
        chalk.green(
          `✓ Extracted all OpenXML parts to ${targetDir}${format ? ' (formatted XML)' : ''}`
        )
      )
      return
    }

    // 2. Extract presentation to raw ZIP archive
    if (opts.zip) {
      const targetZip = resolve(opts.zip)
      await ppt.extractToZip(targetZip)
      console.log(chalk.green(`✓ Extracted presentation ZIP archive to ${targetZip}`))
      return
    }

    // 3. Extract specific slide XML
    if (opts.slide) {
      const slideNum = parseInt(opts.slide, 10)
      const zipManager = ppt.zipManager
      const slidePath = `ppt/slides/slide${slideNum}.xml`
      let xml = await zipManager.readFile(slidePath)

      if (!xml) {
        console.error(chalk.red(`Slide ${slideNum} not found`))
        process.exit(1)
      }

      if (opts.format || opts.formatXml) {
        xml = formatXml(xml)
      }

      if (opts.out) {
        await writeFile(resolve(opts.out), xml, 'utf-8')
        console.log(chalk.green(`✓ Extracted slide ${slideNum} to ${opts.out}`))
      } else {
        console.log(xml)
      }
      return
    }

    // 4. Extract chart XML
    if (opts.chart) {
      const chartName = opts.chart
      const zipManager = ppt.zipManager
      const candidatePaths = [
        `ppt/charts/${chartName}.xml`,
        `ppt/charts/${chartName}`,
        `ppt/charts/chart${chartName}.xml`,
      ]

      let chartXml = null
      let foundPath = null
      for (const p of candidatePaths) {
        if (zipManager.hasFile(p)) {
          chartXml = await zipManager.readFile(p)
          foundPath = p
          break
        }
      }

      if (!chartXml) {
        console.error(chalk.red(`Chart "${chartName}" not found`))
        process.exit(1)
      }

      if (opts.format || opts.formatXml) {
        chartXml = formatXml(chartXml)
      }

      if (opts.out) {
        await writeFile(resolve(opts.out), chartXml, 'utf-8')
        console.log(chalk.green(`✓ Extracted chart ${foundPath} to ${opts.out}`))
      } else {
        console.log(chartXml)
      }
      return
    }

    // 5. Extract relationships
    if (opts.rels) {
      const zipManager = ppt.zipManager
      const relsXml = (await zipManager.readFile('_rels/.rels')) || ''
      if (opts.out) {
        await writeFile(resolve(opts.out), relsXml, 'utf-8')
        console.log(chalk.green(`✓ Extracted _rels/.rels to ${opts.out}`))
      } else {
        console.log(relsXml)
      }
      return
    }

    console.log(
      chalk.yellow(
        'Specify --slide <number>, --chart <name>, --rels, --all, or --zip <file> to extract'
      )
    )
    process.exit(1)
  } catch (err) {
    console.error(chalk.red(`Extract failed: ${err.message}`))
    process.exit(1)
  }
}

module.exports = { extractCommand }
