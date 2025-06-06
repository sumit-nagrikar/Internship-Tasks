const mongoose = require('mongoose');

const semesterSchema = new mongoose.Schema({
    semCode: {
        type: Number,
        required: true,
        index: true,
    },
    semName: {
        type: String,
        required: true,
        trim: true,
    },
    deptCode: {
        type: String,
        required: true,
        ref: 'Department',
        index: true,
    },
    courseCode: {
        type: String,
        required: true,
        ref: 'Course',
        index: true,
    },
    orgCode: {
        type: Number,
        required: true,
        ref: 'Organization',
        index: true,
    },
    orgName: {
        type: String,
        required: true,
        ref: 'Organization',
    },
}, {
    timestamps: true,
});

// Composite unique index to prevent duplicate semesters within the same department, course, and org
semesterSchema.index({ semCode: 1, deptCode: 1, courseCode: 1, orgCode: 1 });

module.exports = mongoose.model('Semester', semesterSchema);
