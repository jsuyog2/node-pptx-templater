/**
 * @fileoverview `build` CLI command — builds a PPTX from a template + JSON data.
 */

import chalk from 'chalk';
import ora from 'ora';
import { readFileSync } from 'fs';
import { resolve } from 'path';
import { PPTXTemplater } from '../../index.js';

/**
 * Executes the `build` CLI command.
 *
 * @param {string} templatePath - Path to the template PPTX.
 * @param {string} outputPath - Path for the generated PPTX.
 * @param {Object} opts - CLI options.
 */
export async function buildCommand(templatePath, outputPath, opts) {
  const spinner = ora(`Loading template: ${templatePath}`).start();

  try {
    // Load the template
    const ppt = await PPTXTemplater.load(resolve(templatePath));
    spinner.succeed(`Loaded template (${ppt.slideCount} slides)`);

    // Apply slide filter if provided
    if (opts.slide) {
      const slideNumbers = opts.slide.split(',').map(n => parseInt(n.trim(), 10));
      ppt.useSlide(...slideNumbers);
      console.log(chalk.cyan(`  → Using slides: ${slideNumbers.join(', ')}`));
    }

    // Apply data from JSON file if provided
    if (opts.data) {
      spinner.start(`Loading data: ${opts.data}`);
      const data = JSON.parse(readFileSync(resolve(opts.data), 'utf-8'));

      // Apply text replacements
      if (data.text) {
        ppt.replaceText(data.text);
        console.log(chalk.cyan(`  → Replaced ${Object.keys(data.text).length} text placeholder(s)`));
      }

      // Apply chart updates
      if (data.charts) {
        for (const [chartId, chartData] of Object.entries(data.charts)) {
          ppt.updateChart(chartId, chartData);
          console.log(chalk.cyan(`  → Updated chart: ${chartId}`));
        }
      }

      // Apply table updates
      if (data.tables) {
        for (const [tableId, tableData] of Object.entries(data.tables)) {
          ppt.updateTable(tableId, tableData);
          console.log(chalk.cyan(`  → Updated table: ${tableId}`));
        }
      }

      spinner.succeed('Data applied');
    }

    // Save output
    spinner.start(`Saving to: ${outputPath}`);
    await ppt.saveToFile(resolve(outputPath));
    spinner.succeed(chalk.green(`✓ Saved: ${outputPath}`));

    console.log(chalk.dim(`\n  Slides: ${ppt.slideCount}`));
    console.log(chalk.dim(`  Template: ${templatePath}`));
    console.log(chalk.dim(`  Output: ${outputPath}\n`));

  } catch (err) {
    spinner.fail(chalk.red(`Build failed: ${err.message}`));
    if (process.env.PPTX_LOG_LEVEL === 'debug') {
      console.error(err.stack);
    }
    process.exit(1);
  }
}
