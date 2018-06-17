function getArguments (container) {
	let arguments = {};
	let input = process.argv.slice(2,process.argv.length);
	for (let i = 0; i < input.length; i += 2) {
		let inputCommand = input[i].replace(/^-+/, '');
		let command = container.find((element) => element.command === inputCommand || element.alias === inputCommand);
		if (command) {
			if (
				command.command === 'help' ||
				(input[i+1] && ['h','help'].includes(input[i+1].replace(/^-+/, '')))
			){
				arguments = {'help': true};
				break;
			}			
			arguments[command.command] = input[i+1] || true;
		}
	}
	return arguments;
}
module.exports = getArguments;