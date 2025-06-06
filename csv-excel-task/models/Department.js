const mongoose = require('mongoose');

const departmentSchema = new mongoose.Schema({
    deptCode: {
        type: String,
        required: true,
        trim: true,
        index: true, // to speed up queries by deptCode
    },
    deptName: {
        type: String,
        required: true,
        trim: true,
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
        index: true,
    },
}, {
    timestamps: true,
});

// Ensure deptCode is unique within the combination of courseCode and orgCode
departmentSchema.index({ deptCode: 1, courseCode: 1, orgCode: 1 });

module.exports = mongoose.model('Department', departmentSchema);
