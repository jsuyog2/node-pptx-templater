/**
 * @fileoverview `validate` CLI command.
 */
import chalk from 'chalk';
import ora from 'ora';
import { resolve } from 'path';
import { PPTXTemplater } from '../../index.js';

export async function validateCommand(filePath, opts) {
  const spinner = ora(`Validating: ${filePath}`).start();
  try {
    const ppt = await PPTXTemplater.load(resolve(filePath));
    const result = ppt.validate();
    spinner.stop();

    if (result.valid && result.warnings.length === 0) {
      console.log(chalk.green(`\n✓ Valid PPTX (${ppt.slideCount} slides)\n`));
    } else {
      if (result.errors.length > 0) {
        console.log(chalk.red(`\n✗ Validation errors (${result.errors.length}):\n`));
        result.errors.forEach(e => console.log(chalk.red(`  • ${e}`)));
      }
      if (result.warnings.length > 0) {
        console.log(chalk.yellow(`\n⚠ Warnings (${result.warnings.length}):\n`));
        result.warnings.forEach(w => console.log(chalk.yellow(`  • ${w}`)));
      }
    }

    if (!result.valid || (opts.strict && result.warnings.length > 0)) {
      process.exit(1);
    }
  } catch (err) {
    spinner.fail(chalk.red(`Validation failed: ${err.message}`));
    process.exit(1);
  }
}
