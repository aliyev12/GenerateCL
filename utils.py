def handle_args(argv):
    commands = ['-h', '-help', '--help', '--h', '-to', '-a']
    commands_desc = {'-h': 'Display help menu',
                     '-to': 'Add HR Manager (attention) block',
                     '-a': 'Add announcement/control numbers'
                     }
    is_help_menu = False
    from_args = {
        "include_to_block": False,
        "include_anncmnt_num": False
    }
    if len(argv) > 1:
        for arg in argv[1:]:
            if arg in commands:
                # Handle help menu
                if arg in commands[:4]:
                    print('')
                    print("Available commands are:")
                    print(
                        f"\t-h, -help, --help, or --h => {commands_desc['-h']}")
                    for command in commands[4:]:
                        print(
                            f"\t{command} => {commands_desc[command]}")
                    print('')
                    is_help_menu = True
                elif arg == '-to':
                    from_args['include_to_block'] = True
                elif arg == '-a':
                    from_args['include_anncmnt_num'] = True

    return (from_args, is_help_menu)
