const fs = require('fs')
const https = require('https')
const reg = require('readable-regex')
const sb = require('structure-bytes')
const xlsx = require('xlsx')

const RANGE_MATCH = reg([
	reg.START,
	'A1:',
	reg.some(
		reg.charIn(['A', 'Z'])
	),
	reg.capture(
		reg.some(
			reg.charIn(['0', '9'])
		),
		'lastRow'
	)
])
const RED_AUTON_PARTICLES = {
	center: 'AI',
	corner: 'AJ'
}
const RED_TELEOP_PARTICLES = {
	center: 'AN',
	corner: 'AO'
}
const BLUE_AUTON_PARTICLES = {
	center: 'AW',
	corner: 'AX'
}
const BLUE_TELEOP_PARTICLES = {
	center: 'BB',
	corner: 'BC'
}
const MATCH_TEAMS = 2
const COLUMNS = {
	date: 'A',
	event: 'B',
	matchType: 'F',
	matchNumber: 'G',
	redTeams: ['H', 'I', 'J'],
	blueTeams: ['K', 'L', 'M'],
	redStatuses: ['N', 'O', 'P'],
	blueStatuses: ['T', 'U', 'V'],
	redScore: {
		autonomousBeacons: 'AG',
		capBallRemoved: 'AH',
		autonomousParticles: RED_AUTON_PARTICLES,
		placements: ['AK', 'AL'],
		teleopBeacons: 'AM',
		teleopParticles: RED_TELEOP_PARTICLES,
		capBallPlacement: 'AP',
		minorPenalties: 'AQ', //caused by this alliance
		majorPenalties: 'AR'
	},
	blueScore: {
		autonomousBeacons: 'AU',
		capBallRemoved: 'AV',
		autonomousParticles: BLUE_AUTON_PARTICLES,
		placements: ['AY', 'AZ'],
		teleopBeacons: 'BA',
		teleopParticles: BLUE_TELEOP_PARTICLES,
		capBallPlacement: 'BD',
		minorPenalties: 'BE', //caused by this alliance
		majorPenalties: 'BF'
	}
}
const QUALIFICATION = 'QUALIFICATION',
	SEMIFINAL = 'SEMIFINAL',
	FINAL = 'FINAL'
function getMatchType(typeString) {
	switch (typeString) {
		case '1':
			return QUALIFICATION
		case '3':
			return SEMIFINAL
		case '4':
			return FINAL
		default:
			throw new Error('No such match type: ' + typeString)
	}
}
const STATUSES = ['ATTENDED', 'NO_SHOW', 'DISQUALIFIED']
const teamType = new sb.StructType({
	number: new sb.UnsignedShortType,
	status: new sb.EnumType({
		type: new sb.StringType,
		values: STATUSES
	})
})
const teamsType = new sb.ChoiceType([
	new sb.TupleType({
		type: teamType,
		length: MATCH_TEAMS
	}),
	new sb.TupleType({
		type: teamType,
		length: MATCH_TEAMS + 1
	})
])
const particleCountType = new sb.StructType({
	center: new sb.UnsignedByteType,
	corner: new sb.UnsignedByteType
})
const ZONES = [null, 'CENTER_HALF', 'CENTER_FULL', 'CORNER_HALF', 'CORNER_FULL']
const CAP_BALL_ZONES = [null, 'LOW', 'HIGH', 'CAPPED']
const scoreType = new sb.StructType({
	auton: new sb.StructType({
		beacons: new sb.UnsignedByteType,
		capBall: new sb.BooleanType,
		particles: particleCountType,
		placements: new sb.TupleType({
			type: new sb.EnumType({
				type: new sb.OptionalType(new sb.StringType),
				values: ZONES
			}),
			length: MATCH_TEAMS
		})
	}),
	teleop: new sb.StructType({
		beacons: new sb.UnsignedByteType,
		particles: particleCountType,
		capBall: new sb.EnumType({
			type: new sb.OptionalType(new sb.StringType),
			values: CAP_BALL_ZONES
		})
	}),
	penalties: new sb.StructType({
		minor: new sb.UnsignedByteType,
		major: new sb.UnsignedByteType
	})
})
const type = new sb.MapType(
	new sb.StructType({
		month: new sb.UnsignedByteType,
		day: new sb.UnsignedByteType,
		name: new sb.StringType
	}),
	new sb.ArrayType(
		new sb.StructType({
			type: new sb.EnumType({
				type: new sb.StringType,
				values: [QUALIFICATION, SEMIFINAL, FINAL]
			}),
			number: new sb.UnsignedByteType,
			redTeams: teamsType,
			blueTeams: teamsType,
			redScore: scoreType,
			blueScore: scoreType
		})
	)
)
const SPACE = ' ', ZERO = '0'
https.get('https://standings.firstinspires.org/ftc/Scoring-System-Results.xlsx', res => {
	const chunks = []
	res
		.on('error', err => {
			throw err
		})
		.on('data', chunk => chunks.push(chunk))
		.on('end', () => {
			parseFile(Buffer.concat(chunks))
		})
})
function parseFile(data) {
	const document = xlsx.read(data, {
		cellDates: true,
		cellFormula: false,
		cellHTML: false
	})
	const sheet = document.Sheets.Sheet1
	const range = sheet['!ref']
	const rangeMatch = reg.exec(RANGE_MATCH, range)
	const lastRow = Number(rangeMatch.get('lastRow'))
	const results = new Map
	const events = {}
	for (let row = 2; row < lastRow; row++) {
		const rowString = String(row)
		const typeString = sheet[COLUMNS.matchType + rowString].w
		if (typeString === ZERO) continue //practice match?
		const eventName = sheet[COLUMNS.event + rowString].w
		let event = events[eventName]
		if (!event) {
			const dateCell = sheet[COLUMNS.date + rowString].w
			const dateSpaceIndex = dateCell.indexOf(SPACE)
			let dateString
			if (dateSpaceIndex === -1) dateString = dateCell
			else dateString = dateCell.substring(0, dateSpaceIndex)
			const date = new Date(dateString)
			event = events[eventName] = {
				month: date.getMonth() + 1,
				day: date.getDate(),
				name: eventName
			}
			results.set(event, [])
		}
		const redTeams = []
		for (let i = 0; i < COLUMNS.redTeams.length; i++) {
			const number = sheet[COLUMNS.redTeams[i] + rowString].v
			if (!number) break
			const status = STATUSES[sheet[COLUMNS.redStatuses[i] + rowString].v]
			redTeams[i] = {number, status}
		}
		const blueTeams = []
		for (let i = 0; i < COLUMNS.blueTeams.length; i++) {
			const number = sheet[COLUMNS.blueTeams[i] + rowString].v
			if (!number) break
			const status = STATUSES[sheet[COLUMNS.blueStatuses[i] + rowString].w]
			blueTeams[i] = {number, status}
		}
		const redPlacements = new Array(MATCH_TEAMS)
		for (let i = 0; i < redPlacements.length; i++) {
			redPlacements[i] = ZONES[sheet[COLUMNS.redScore.placements[i] + rowString].v]
		}
		const redAutonParticles = {}
		for (const type in RED_AUTON_PARTICLES) {
			redAutonParticles[type] = sheet[RED_AUTON_PARTICLES[type] + rowString].v
		}
		const redTeleopParticles = {}
		for (const type in RED_TELEOP_PARTICLES) {
			redTeleopParticles[type] = sheet[RED_TELEOP_PARTICLES[type] + rowString].v
		}
		const redScore = {
			auton: {
				beacons: sheet[COLUMNS.redScore.autonomousBeacons + rowString].v,
				capBall: sheet[COLUMNS.redScore.capBallRemoved + rowString].v,
				particles: redAutonParticles,
				placements: redPlacements
			},
			teleop: {
				beacons: sheet[COLUMNS.redScore.teleopBeacons + rowString].v,
				particles: redTeleopParticles,
				capBall: CAP_BALL_ZONES[sheet[COLUMNS.redScore.capBallPlacement + rowString].v]
			},
			penalties: {
				minor: sheet[COLUMNS.redScore.minorPenalties + rowString].v,
				major: sheet[COLUMNS.redScore.majorPenalties + rowString].v
			}
		}
		const bluePlacements = new Array(MATCH_TEAMS)
		for (let i = 0; i < bluePlacements.length; i++) {
			bluePlacements[i] = ZONES[sheet[COLUMNS.blueScore.placements[i] + rowString].v]
		}
		const blueAutonParticles = {}
		for (const type in BLUE_AUTON_PARTICLES) {
			blueAutonParticles[type] = sheet[BLUE_AUTON_PARTICLES[type] + rowString].v
		}
		const blueTeleopParticles = {}
		for (const type in BLUE_TELEOP_PARTICLES) {
			blueTeleopParticles[type] = sheet[BLUE_TELEOP_PARTICLES[type] + rowString].v
		}
		const blueScore = {
			auton: {
				beacons: sheet[COLUMNS.blueScore.autonomousBeacons + rowString].v,
				capBall: sheet[COLUMNS.blueScore.capBallRemoved + rowString].v,
				particles: blueAutonParticles,
				placements: bluePlacements
			},
			teleop: {
				beacons: sheet[COLUMNS.blueScore.teleopBeacons + rowString].v,
				particles: blueTeleopParticles,
				capBall: CAP_BALL_ZONES[sheet[COLUMNS.blueScore.capBallPlacement + rowString].v]
			},
			penalties: {
				minor: sheet[COLUMNS.blueScore.minorPenalties + rowString].v,
				major: sheet[COLUMNS.blueScore.majorPenalties + rowString].v
			}
		}
		results.get(event).push({
			type: getMatchType(typeString),
			number: sheet[COLUMNS.matchNumber + rowString].v,
			redTeams,
			blueTeams,
			redScore,
			blueScore
		})
	}
	sb.writeTypeAndValue({
		type,
		value: results,
		outStream: fs.createWriteStream(__dirname + '/results.sbtv')
	}, err => {
		if (err) throw err
	})
}