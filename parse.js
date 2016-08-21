const fs = require('fs');
const sb = require('structure-bytes');
const xlsx = require('xlsx');

const RANGE_MATCH = /^A1:([A-Z]+)([0-9]+)$/;
const COLUMNS = {
	date: 'A',
	event: 'B',
	matchType: 'F',
	matchNumber: 'G',
	redTeams: ['H', 'I', 'J'],
	blueTeams: ['K', 'L', 'M'],
	redStatuses: ['N', 'O', 'P'],
	blueStatuses: ['Q', 'R', 'S'],
	redScore: {
		autonomousPlacements: ['AA', 'AB'],
		rescueBeacons: 'AC',
		autonomousClimbers: 'AD',
		teleopPlacements: ['AE', 'AF'],
		floorGoal: 'AG',
		highGoal: 'AH',
		lowGoal: 'AI',
		midGoal: 'AJ',
		teleopClimbers: 'AK', //includes autonomous ones - there are 356 instances where (autonClimbers + teleopClimbers > 4) and only 33 where (autonClimbers > teleopClimbers), which suggests that they were usually including autonomous ones in the teleop count
		zips: 'AL',
		allClears: 'AM',
		hanging: 'AN',
		minorPenalties: 'AO', //caused by this alliance
		majorPenalties: 'AP'
	},
	blueScore: {
		autonomousPlacements: ['AS', 'AT'],
		rescueBeacons: 'AU',
		autonomousClimbers: 'AV',
		teleopPlacements: ['AW', 'AX'],
		floorGoal: 'AY',
		highGoal: 'AZ',
		lowGoal: 'BA',
		midGoal: 'BB',
		teleopClimbers: 'BC', //includes autonomous ones
		zips: 'BD',
		allClears: 'BE',
		hanging: 'BF',
		minorPenalties: 'BG', //caused by this alliance
		majorPenalties: 'BH'
	}
};
const QUALIFICATION = 'QUALIFICATION',
	SEMIFINAL = 'SEMIFINAL',
	FINAL = 'FINAL';
function getMatchType(typeString) {
	switch (typeString) {
		case '1':
			return QUALIFICATION;
		case '3':
			return SEMIFINAL;
		case '4':
			return FINAL;
		default:
			throw new Error('No such match type: ' + typeString);
	}
}
const ATTENDED = 'ATTENDED',
	NO_SHOW = 'NO_SHOW',
	DISQUALIFIED = 'DISQUALIFIED';
function getStatus(statusString) {
	switch (statusString) {
		case '0':
			return ATTENDED;
		case '1':
			return NO_SHOW;
		case '2':
			return DISQUALIFIED;
		default:
			throw new Error('No such status: ' + statusString);
	}
}
const REPAIR_ZONE = 'REPAIR_ZONE',
	FLOOR_GOAL = 'FLOOR_GOAL',
	MOUNTAIN_TOUCHING_FLOOR = 'HALF_ON',
	LOW_ZONE = 'LOW_ZONE',
	MID_ZONE = 'MID_ZONE',
	HIGH_ZONE = 'HIGH_ZONE';
const ZONES = [null, REPAIR_ZONE, FLOOR_GOAL, MOUNTAIN_TOUCHING_FLOOR, LOW_ZONE, MID_ZONE, HIGH_ZONE];
const teamType = new sb.StructType({
	number: new sb.UnsignedShortType,
	status: new sb.EnumType({
		type: new sb.StringType,
		values: [ATTENDED, NO_SHOW, DISQUALIFIED]
	})
});
const teamsType = new sb.ChoiceType([
	new sb.TupleType({
		type: teamType,
		length: 1
	}),
	new sb.TupleType({
		type: teamType,
		length: 2
	}),
	new sb.TupleType({
		type: teamType,
		length: 3
	})
]);
const placementType = new sb.TupleType({
	type: new sb.EnumType({
		type: new sb.OptionalType(new sb.StringType),
		values: ZONES
	}),
	length: 2
});
const scoreType = new sb.StructType({
	auton: new sb.StructType({
		placements: placementType,
		beacons: new sb.UnsignedByteType,
		climbers: new sb.UnsignedByteType
	}),
	teleop: new sb.StructType({
		placements: placementType,
		floor: new sb.UnsignedByteType,
		low: new sb.UnsignedByteType,
		mid: new sb.UnsignedByteType,
		high: new sb.UnsignedByteType,
		climbers: new sb.UnsignedByteType, //not including autonomous ones
		zips: new sb.UnsignedByteType,
		allClears: new sb.UnsignedByteType,
		hanging: new sb.UnsignedByteType
	}),
	penalties: new sb.StructType({
		minor: new sb.UnsignedByteType,
		major: new sb.UnsignedByteType
	})
});
const type = new sb.ArrayType(
	new sb.StructType({
		month: new sb.UnsignedByteType,
		day: new sb.UnsignedByteType,
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
);
const SPACE = ' ';
fs.readFile(__dirname + '/Scoring-System-Results.xlsx', (err, data) => {
	if (err) throw err;
	const document = xlsx.read(data, {
		cellDates: true,
		cellFormula: false,
		cellHTML: false
	});
	const sheet = document.Sheets.Sheet1;
	const range = sheet['!ref'];
	const rangeMatch = RANGE_MATCH.exec(range);
	const lastRow = Number(rangeMatch[2]);
	const redAutonPlacementColumns = COLUMNS.redScore.autonomousPlacements,
		redTeleopPlacementColumns = COLUMNS.redScore.teleopPlacements,
		blueAutonPlacementColumns = COLUMNS.blueScore.autonomousPlacements,
		blueTeleopPlacementColumns = COLUMNS.blueScore.teleopPlacements;
	const results = [];
	for (let row = 2; row < lastRow; row++) {
		const rowString = String(row);
		const dateCell = sheet[COLUMNS.date + rowString].w;
		const dateSpaceIndex = dateCell.indexOf(SPACE);
		let dateString;
		if (dateSpaceIndex === -1) dateString = dateCell;
		else dateString = dateCell.substring(0, dateSpaceIndex);
		const date = new Date(dateString);
		const typeString = sheet[COLUMNS.matchType + rowString].w;
		if (typeString === '0') continue; //practice match?
		const redTeams = [];
		for (let i = 0; i < COLUMNS.redTeams.length; i++) {
			const teamNumber = sheet[COLUMNS.redTeams[i] + rowString].v;
			if (!teamNumber) break;
			const status = getStatus(sheet[COLUMNS.redStatuses[i] + rowString].w);
			redTeams[i] = {number: teamNumber, status};
		}
		const blueTeams = [];
		for (let i = 0; i < COLUMNS.blueTeams.length; i++) {
			const teamNumber = sheet[COLUMNS.blueTeams[i] + rowString].v;
			if (!teamNumber) break;
			const status = getStatus(sheet[COLUMNS.blueStatuses[i] + rowString].w);
			blueTeams[i] = {number: teamNumber, status};
		}
		const redAutonPlacements = new Array(2);
		for (let i = 0; i < redAutonPlacementColumns.length; i++) {
			redAutonPlacements[i] = ZONES[sheet[redAutonPlacementColumns[i] + rowString].v];
		}
		const redTeleopPlacements = new Array(2);
		for (let i = 0; i < redTeleopPlacementColumns.length; i++) {
			redTeleopPlacements[i] = ZONES[sheet[redTeleopPlacementColumns[i] + rowString].v];
		}
		const redAutonClimbers = sheet[COLUMNS.redScore.autonomousClimbers + rowString].v;
		const redScore = {
			auton: {
				placements: redAutonPlacements,
				beacons: sheet[COLUMNS.redScore.rescueBeacons + rowString].v,
				climbers: redAutonClimbers
			},
			teleop: {
				placements: redTeleopPlacements,
				floor: sheet[COLUMNS.redScore.floorGoal + rowString].v,
				low: sheet[COLUMNS.redScore.lowGoal + rowString].v,
				mid: sheet[COLUMNS.redScore.midGoal + rowString].v,
				high: sheet[COLUMNS.redScore.highGoal + rowString].v,
				climbers: Math.max(sheet[COLUMNS.redScore.teleopClimbers + rowString].v - redAutonClimbers, 0),
				zips: sheet[COLUMNS.redScore.zips + rowString].v,
				allClears: sheet[COLUMNS.redScore.allClears + rowString].v,
				hanging: sheet[COLUMNS.redScore.hanging + rowString].v
			},
			penalties: {
				minor: sheet[COLUMNS.redScore.minorPenalties + rowString].v,
				major: sheet[COLUMNS.redScore.majorPenalties + rowString].v
			}
		};
		const blueAutonPlacements = new Array(2);
		for (let i = 0; i < blueAutonPlacementColumns.length; i++) {
			blueAutonPlacements[i] = ZONES[sheet[blueAutonPlacementColumns[i] + rowString].v];
		}
		const blueTeleopPlacements = new Array(2);
		for (let i = 0; i < blueTeleopPlacementColumns.length; i++) {
			blueTeleopPlacements[i] = ZONES[sheet[blueTeleopPlacementColumns[i] + rowString].v];
		}
		const blueAutonClimbers = sheet[COLUMNS.blueScore.autonomousClimbers + rowString].v;
		const blueScore = {
			auton: {
				placements: blueAutonPlacements,
				beacons: sheet[COLUMNS.blueScore.rescueBeacons + rowString].v,
				climbers: blueAutonClimbers
			},
			teleop: {
				placements: blueTeleopPlacements,
				floor: sheet[COLUMNS.blueScore.floorGoal + rowString].v,
				low: sheet[COLUMNS.blueScore.lowGoal + rowString].v,
				mid: sheet[COLUMNS.blueScore.midGoal + rowString].v,
				high: sheet[COLUMNS.blueScore.highGoal + rowString].v,
				climbers: Math.max(sheet[COLUMNS.blueScore.teleopClimbers + rowString].v - blueAutonClimbers, 0),
				zips: sheet[COLUMNS.blueScore.zips + rowString].v,
				allClears: sheet[COLUMNS.blueScore.allClears + rowString].v,
				hanging: sheet[COLUMNS.blueScore.hanging + rowString].v
			},
			penalties: {
				minor: sheet[COLUMNS.blueScore.minorPenalties + rowString].v,
				major: sheet[COLUMNS.blueScore.majorPenalties + rowString].v
			}
		};
		results.push({
			month: date.getMonth() + 1,
			day: date.getDate(),
			type: getMatchType(typeString),
			number: sheet[COLUMNS.matchNumber + rowString].v,
			redTeams,
			blueTeams,
			redScore,
			blueScore
		});
	}
	sb.writeTypeAndValue({
		type,
		value: results,
		outStream: fs.createWriteStream(__dirname + '/results.sbtv')
	}, err => {
		if (err) throw err;
	});
});