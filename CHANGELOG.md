# Change Log
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [Unreleased]


## [v1.2.0] - 2018-01-11
### Added
- Return all outputs of a script with delimiter from App.config

### Fixed
- Fix bug when script has multiple outputs


## [v1.1.3] - 2018-01-09
### Fixed
- Fix bug with concurrent script execution when using message queue
- Fix bug with empty message body or invalid JSON


## [v1.1.2] - 2017-12-12
### Changed
- Refactor PSScriptExecutor methods

### Fixed
- Fix issue with default max runspaces (only 1)


## [v1.1.1] - 2017-08-30
### Changed
- RabbitMQ response content type to text/plain
- Add double quotes to param values (needed for spaces etc.)


## [v1.1.0] - 2017-08-24
### Added
- RabbitMQ message consumer for asynchronous script execution
- HTTP: Add possibility to send parameters as a JSON in the request body

### Changed
- Refactor code: extract modules into separate classes


## [v1.0.0] - 2017-08-22
### Added
- First release version


[Unreleased]: https://github.com/dwettstein/PSScriptInvoker/compare/v1.2.0...HEAD
[v1.2.0]: https://github.com/dwettstein/PSScriptInvoker/compare/v1.1.3...v1.2.0
[v1.1.3]: https://github.com/dwettstein/PSScriptInvoker/compare/v1.1.2...v1.1.3
[v1.1.2]: https://github.com/dwettstein/PSScriptInvoker/compare/v1.1.1...v1.1.2
[v1.1.1]: https://github.com/dwettstein/PSScriptInvoker/compare/v1.1.0...v1.1.1
[v1.1.0]: https://github.com/dwettstein/PSScriptInvoker/compare/v1.0.0...v1.1.0
[v1.0.0]: https://github.com/dwettstein/PSScriptInvoker/tree/v1.0.0
