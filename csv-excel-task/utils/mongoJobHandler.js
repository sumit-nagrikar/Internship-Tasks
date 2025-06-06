const mongoose = require('mongoose');
const Organization = require('../models/Organization');
const Course = require('../models/Course');
const Department = require('../models/Department');
const Semester = require('../models/Semester');
const Student = require('../models/Student');

// Each entity (course, department, etc.) has some key fields that are used to identify uniqueness.
const entityKeys = {
    courses: ['COURSECODE'],
    departments: ['DEPTCODE', 'COURSECODE'],
    semesters: ['SEMCODE', 'DEPTCODE', 'COURSECODE'],
    students: ['NAME', 'ORGCODE', 'SEMCODE', 'COURSECODE', 'DEPTCODE'],
};

// Removes duplicate rows based on specific fields - return only unique rows
function groupByKeys(data, keys, entityName) {
    const map = new Map();
    const delimiter = '|||';
    let skippedRows = 0;

    for (const row of data) {
        if (keys.every(k => k in row && row[k] != null && row[k] !== '')) {
            const key = keys
                .map(k => String(row[k]).trim().toUpperCase().replace(delimiter, ''))
                .join(delimiter);
            if (!map.has(key)) map.set(key, row);
        } else {
            skippedRows++;
        }
    }

    if (skippedRows > 0) {
        console.warn(`Skipped ${skippedRows} rows for ${entityName} due to missing or empty keys: ${keys.join(', ')}`);
    }

    return [...map.values()];
}

// Process data in chunks to handle large datasets
async function processInChunks(data, chunkSize, keys, entityName) {
    const chunks = [];
    for (let i = 0; i < data.length; i += chunkSize) {
        chunks.push(groupByKeys(data.slice(i, i + chunkSize), keys, entityName));
    }
    return chunks.flat();
}


async function processMongoJob({ orgData, orgName, sheetId }) {
    if (!orgData?.length) throw new Error('No org data');
    if (!orgName || orgName === 'ORGNAME') {
        throw new Error(`Invalid orgName: ${orgName}`);
    }

    const session = await mongoose.startSession();
    session.startTransaction();
        console.log("orgData",orgData);
    try {
        // Validate and extract orgCode
        const orgCode = Number(orgData[0]?.ORGCODE);
        console.log("orgCode",orgCode);
        if (isNaN(orgCode)) throw new Error('Invalid ORGCODE in first row');

        // Update/insert Organization
        await Organization.updateOne(
            { orgCode },
            { $set: { orgCode, orgName } },
            { upsert: true, session }
        );

        // Process courses
        const courses = await processInChunks(orgData, 2, entityKeys.courses, 'courses');
    
        const courseOps = courses.map(row => ({
            updateOne: {
                filter: { courseCode: row.COURSECODE, orgName },
                update: { $set: { courseName: row.COURSENAME || '', orgName } },
                upsert: true,
            },
        }));

        if (courseOps.length) await Course.bulkWrite(courseOps, { session });

        // Process departments
        const departments = await processInChunks(orgData, 1000, entityKeys.departments, 'departments');
        const deptOps = departments.map(row => ({
            updateOne: {
                filter: { deptCode: row.DEPTCODE, courseCode: row.COURSECODE, orgName },
                update: { $set: { deptName: row.DEPTNAME || '', courseCode: row.COURSECODE, orgName } },
                upsert: true,
            },
        }));

        if (deptOps.length) await Department.bulkWrite(deptOps, { session });

        // Process semesters
        const semesters = await processInChunks(orgData, 1000, entityKeys.semesters, 'semesters');
        const semOps = semesters.map(row => ({
            updateOne: {
                filter: {
                    semCode: Number(row.SEMCODE),
                    deptCode: row.DEPTCODE,
                    courseCode: row.COURSECODE,
                    orgName,
                },
                update: {
                    $set: {
                        semName: row.SEMNAME || '',
                        deptCode: row.DEPTCODE,
                        courseCode: row.COURSECODE,
                        orgName,
                    },
                },
                upsert: true,
            },
        }));
        if (semOps.length) await Semester.bulkWrite(semOps, { session });

        // Process students with optimized key filtering
        const excludedKeys = new Set([
            'ORGCODE', 'ORGNAME', 'COURSECODE', 'COURSENAME',
            'DEPTCODE', 'DEPTNAME', 'SEMCODE', 'SEMNAME', 'NAME'
        ]);
        const requiredStudentKeys = ['NAME', 'ORGCODE', 'COURSECODE', 'DEPTCODE', 'SEMCODE'];
        const seenStudents = new Set();
        let duplicateStudents = 0;

        const studentDocs = orgData
            .map(row => {
                // Validate required fields/skip row is required fileds missing
                if (requiredStudentKeys.some(k => !(k in row) || row[k] == null || row[k] === '')) {
                    console.warn(`Skipping invalid student row: missing or empty keys ${requiredStudentKeys.join(', ')}`);
                    return null;
                }
                //generate a unique student key
                const studentKey = entityKeys.students
                    .map(k => String(row[k]).trim().toUpperCase())
                    .join('|||');
                if (seenStudents.has(studentKey)) {
                    duplicateStudents++;
                    console.warn(`Duplicate student detected: ${studentKey}`);
                } else {
                    seenStudents.add(studentKey);
                }

                return {
                    name: row.NAME,
                    orgCode,
                    courseCode: row.COURSECODE,
                    deptCode: row.DEPTCODE,
                    semCode: Number(row.SEMCODE),
                    extraFields: Object.fromEntries(
                        Object.entries(row)
                            .filter(([key]) => !excludedKeys.has(key))
                            .map(([k, v]) => [k.toLowerCase(), v ?? null])
                    ),
                };
            })
            .filter(student => student !== null);

        if (duplicateStudents > 0) {
            console.warn(`Detected ${duplicateStudents} potential duplicate students`);
        }

        if (studentDocs.length) await Student.insertMany(studentDocs, { session });

        await session.commitTransaction();
        console.log(`Processed ${orgName} (orgCode: ${orgCode}) with ${studentDocs.length} students for sheet ${sheetId}`);
    } catch (error) {
        await session.abortTransaction();
        try {
            await mongoose.connection.db.collection('error_logs').insertOne({
                orgName,
                sheetId,
                error: error.message,
                stack: error.stack,
                timestamp: new Date(),
            }, { session });
        } catch (logError) {
            console.error('Failed to log error:', logError);
        }
        console.error(`Error processing MongoDB job ${orgName} (sheet ${sheetId}):`, error);
        throw error;
    } finally {
        await session.endSession();
    }
}

module.exports = { processMongoJob };