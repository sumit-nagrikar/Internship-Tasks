const mongoose = require('mongoose');
const courseSchema = new mongoose.Schema({
    courseCode: { type: String, required: true },
    courseName: { type: String, required: true },
    orgName: { ref: 'Organization', type: String, required: true },
},);
courseSchema.index({ courseCode: 1, orgName: 1 });
module.exports = mongoose.model('Course', courseSchema);