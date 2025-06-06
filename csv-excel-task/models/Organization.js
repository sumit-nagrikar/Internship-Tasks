const mongoose = require('mongoose');

const organizationSchema = new mongoose.Schema({
    orgCode: { type: Number, required: true, unique: true },
    orgName: { type: String, required: true },
});

organizationSchema.index({ orgCode: 1 });

module.exports = mongoose.models.Organization || mongoose.model('Organization', organizationSchema);