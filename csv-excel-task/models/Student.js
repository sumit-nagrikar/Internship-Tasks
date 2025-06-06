const mongoose = require('mongoose');

const studentSchema = new mongoose.Schema({
    name: {
        type: String,
        required: true,
        trim: true,
    },
    orgCode: {
        type: Number,
        required: true,
        ref: 'Organization',
        index: true,
    },
    courseCode: {
        type: String,
        required: true,
        ref: 'Course',
        index: true,
    },
    deptCode: {
        type: String,
        required: true,
        ref: 'Department',
        index: true,
    },
    semCode: {
        type: Number,
        required: true,
        ref: 'Semester',
        index: true,
    },
    extraFields: {
        type: mongoose.Schema.Types.Mixed,
    },
}, {
    timestamps: true,
});

// Composite unique index to avoid duplicate student entries within the same org/course/department/semester
studentSchema.index({ name: 1, orgCode: 1, courseCode: 1, deptCode: 1, semCode: 1 });

module.exports = mongoose.model('Student', studentSchema);
