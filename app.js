import express from "express";
import mongoose from "mongoose";
import cors from "cors";
import ExcelJS from "exceljs";
import { config } from "dotenv";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import morgan from "morgan";
import fs from "fs";
import path from "path";

config();

const app = express();
const port = process.env.PORT || 8000;
const secureKey = process.env.SECUREKEY;
const DBurl = process.env.DBURL;
const NODE_ENV = process.env.NODE_ENV || "development";

// Security Middleware
app.use(
  helmet({
    crossOriginResourcePolicy: { policy: "cross-origin" },
  })
);

// Rate Limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: NODE_ENV === "production" ? 100 : 1000,
  message: {
    success: false,
    message: "Too many requests from this IP, please try again later.",
  },
  standardHeaders: true,
  legacyHeaders: false,
});
app.use(limiter);

// Logging Middleware
if (NODE_ENV === "development") {
  app.use(morgan("dev"));
} else {
  const logsDir = path.join(process.cwd(), "logs");
  if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
  }
  const accessLogStream = fs.createWriteStream(
    path.join(logsDir, "access.log"),
    { flags: "a" }
  );
  app.use(morgan("combined", { stream: accessLogStream }));
}

// CORS Configuration
const corsOptions = {
  origin: function (origin, callback) {
    if (!origin) return callback(null, true);

    const allowedOrigins = [
      "http://localhost:3000",
      "http://localhost:5173",
      "https://reunir1.netlify.app/",
      "https://reunir1.netlify.app/admin",
    ];

    if (allowedOrigins.indexOf(origin) !== -1 || NODE_ENV === "development") {
      callback(null, true);
    } else {
      callback(new Error("Not allowed by CORS"));
    }
  },
  credentials: true,
  optionsSuccessStatus: 200,
};
app.use(cors(corsOptions));

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

// Database Connection with retry logic
const DBCon = async (retries = 5, delay = 5000) => {
  for (let i = 0; i < retries; i++) {
    try {
      await mongoose.connect(DBurl, {
        serverSelectionTimeoutMS: 5000,
        socketTimeoutMS: 45000,
      });
      console.log("✅ Database connected successfully");
      return;
    } catch (error) {
      console.error(
        `❌ Database connection attempt ${i + 1} failed:`,
        error.message
      );

      if (i === retries - 1) {
        console.error("💥 All database connection attempts failed");
        process.exit(1);
      }

      console.log(`🔄 Retrying in ${delay / 1000} seconds...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
};

// Health Check Endpoint
app.get("/health", (req, res) => {
  res.status(200).json({
    success: true,
    message: "Server is healthy",
    timestamp: new Date().toISOString(),
    environment: NODE_ENV,
    database:
      mongoose.connection.readyState === 1 ? "connected" : "disconnected",
  });
});

// All Schema Design
const contactSchema = new mongoose.Schema(
  {
    name: {
      type: String,
      required: [true, "Name is required"],
      trim: true,
      maxlength: [100, "Name cannot exceed 100 characters"],
    },
    email: {
      type: String,
      required: [true, "Email is required"],
      trim: true,
      lowercase: true,
      match: [
        /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/,
        "Please enter a valid email",
      ],
    },
    number: {
      type: String,
      required: [true, "Phone number is required"],
      trim: true,
      match: [/^[0-9]{10}$/, "Please enter a valid 10-digit phone number"],
    },
    message: {
      type: String,
      required: [true, "Message is required"],
      trim: true,
      maxlength: [1000, "Message cannot exceed 1000 characters"],
    },
  },
  { timestamps: true }
);

const memberSchema = new mongoose.Schema(
  {
    aadharnumber: {
      type: String,
      required: [true, "Aadhar number is required"],
      unique: true,
      trim: true,
      minlength: [12, "Aadhar number must be 12 digits"],
      maxlength: [12, "Aadhar number must be 12 digits"],
    },
    email: {
      type: String,
      required: [true, "Email is required"],
      unique: true,
      trim: true,
      lowercase: true,
      match: [
        /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/,
        "Please enter a valid email",
      ],
    },
    salutation: {
      type: String,
      trim: true,
      enum: ["Mr", "Mrs", "Ms", "Dr", "Prof"],
      default: "Mr",
    },
    name: {
      type: String,
      required: [true, "Name is required"],
      trim: true,
      maxlength: [100, "Name cannot exceed 100 characters"],
    },
    dob: {
      type: Date,
      validate: {
        validator: function (date) {
          return date <= new Date();
        },
        message: "Date of birth cannot be in the future",
      },
    },
    number: {
      type: String,
      required: [true, "Phone number is required"],
      trim: true,
      match: [/^[0-9]{10}$/, "Please enter a valid 10-digit phone number"],
    },
    address: {
      type: String,
      trim: true,
      maxlength: [200, "Address cannot exceed 200 characters"],
    },
    city: {
      type: String,
      trim: true,
      maxlength: [50, "City cannot exceed 50 characters"],
    },
    state: {
      type: String,
      trim: true,
      maxlength: [50, "State cannot exceed 50 characters"],
    },
    country: {
      type: String,
      trim: true,
      default: "India",
      maxlength: [50, "Country cannot exceed 50 characters"],
    },
    pincode: {
      type: String,
      trim: true,
      match: [/^[0-9]{6}$/, "Please enter a valid 6-digit pincode"],
    },
    profession: {
      type: String,
      trim: true,
      maxlength: [50, "Profession cannot exceed 50 characters"],
    },
    gender: {
      type: String,
      trim: true,
      enum: ["Male", "Female", "Other"],
      required: true,
    },
    bloodgroup: {
      type: String,
      trim: true,
      enum: ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-", null],
      uppercase: true,
    },
    emgname: {
      type: String,
      trim: true,
      maxlength: [100, "Emergency contact name cannot exceed 100 characters"],
    },
    emgnumber: {
      type: String,
      trim: true,
      match: [
        /^[0-9]{10}$/,
        "Please enter a valid 10-digit emergency contact number",
      ],
    },
    sportperson: {
      type: Boolean,
      default: false,
    },
  },
  {
    timestamps: true,
  }
);

const athleteSchema = new mongoose.Schema(
  {
    aadharnumber: {
      type: String,
      required: [true, "Aadhar number is required"],
      unique: true,
      trim: true,
      minlength: [12, "Aadhar number must be 12 digits"],
      maxlength: [12, "Aadhar number must be 12 digits"],
    },
    email: {
      type: String,
      required: [true, "Email is required"],
      unique: true,
      trim: true,
      lowercase: true,
      match: [
        /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/,
        "Please enter a valid email",
      ],
    },
    salutation: {
      type: String,
      trim: true,
      enum: ["Mr", "Mrs", "Ms", "Dr", "Prof"],
      default: "Mr",
    },
    name: {
      type: String,
      required: [true, "Name is required"],
      trim: true,
      maxlength: [100, "Name cannot exceed 100 characters"],
    },
    dob: {
      type: Date,
      validate: {
        validator: function (date) {
          return date <= new Date();
        },
        message: "Date of birth cannot be in the future",
      },
    },
    number: {
      type: String,
      required: [true, "Phone number is required"],
      trim: true,
      match: [/^[0-9]{10}$/, "Please enter a valid 10-digit phone number"],
    },
    address: {
      type: String,
      trim: true,
      maxlength: [200, "Address cannot exceed 200 characters"],
    },
    city: {
      type: String,
      trim: true,
      maxlength: [50, "City cannot exceed 50 characters"],
    },
    state: {
      type: String,
      trim: true,
      maxlength: [50, "State cannot exceed 50 characters"],
    },
    country: {
      type: String,
      trim: true,
      default: "India",
      maxlength: [50, "Country cannot exceed 50 characters"],
    },
    pincode: {
      type: String,
      trim: true,
      match: [/^[0-9]{6}$/, "Please enter a valid 6-digit pincode"],
    },
    profession: {
      type: String,
      trim: true,
      maxlength: [50, "Profession cannot exceed 50 characters"],
    },
    gender: {
      type: String,
      trim: true,
      enum: ["Male", "Female", "Other"],
      required: true,
    },
    bloodgroup: {
      type: String,
      trim: true,
      enum: ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-", null],
      uppercase: true,
    },
    emgname: {
      type: String,
      trim: true,
      maxlength: [100, "Emergency contact name cannot exceed 100 characters"],
    },
    emgnumber: {
      type: String,
      trim: true,
      match: [
        /^[0-9]{10}$/,
        "Please enter a valid 10-digit emergency contact number",
      ],
    },
    sportperson: {
      type: Boolean,
      default: true,
    },
    sportDiscipline: {
      type: String,
      required: [true, "Sport discipline is required"],
      trim: true,
      maxlength: [100, "Sport discipline cannot exceed 100 characters"],
    },
    levelOfSport: {
      type: String,
      required: [true, "Level of sport is required"],
      trim: true,
      enum: ["Beginner", "Intermediate", "Advanced", "Professional", "Elite"],
    },
    sportsTitleEarned: {
      type: String,
      trim: true,
      maxlength: [200, "Sports title cannot exceed 200 characters"],
    },
    club: {
      type: String,
      trim: true,
      maxlength: [100, "Club name cannot exceed 100 characters"],
    },
    achievements: {
      type: String,
      maxlength: [500, "Achievements cannot exceed 500 characters"],
    },
  },
  {
    timestamps: true,
  }
);

const coachSchema = new mongoose.Schema(
  {
    aadharnumber: {
      type: String,
      required: [true, "Aadhar number is required"],
      unique: true,
      trim: true,
      minlength: [12, "Aadhar number must be 12 digits"],
      maxlength: [12, "Aadhar number must be 12 digits"],
    },
    email: {
      type: String,
      required: [true, "Email is required"],
      unique: true,
      trim: true,
      lowercase: true,
      match: [
        /^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/,
        "Please enter a valid email",
      ],
    },
    salutation: {
      type: String,
      trim: true,
      enum: ["Mr", "Mrs", "Ms", "Dr", "Prof"],
      default: "Mr",
    },
    name: {
      type: String,
      required: [true, "Name is required"],
      trim: true,
      maxlength: [100, "Name cannot exceed 100 characters"],
    },
    dob: {
      type: Date,
      validate: {
        validator: function (date) {
          return date <= new Date();
        },
        message: "Date of birth cannot be in the future",
      },
    },
    number: {
      type: String,
      required: [true, "Phone number is required"],
      trim: true,
      match: [/^[0-9]{10}$/, "Please enter a valid 10-digit phone number"],
    },
    address: {
      type: String,
      trim: true,
      maxlength: [200, "Address cannot exceed 200 characters"],
    },
    city: {
      type: String,
      trim: true,
      maxlength: [50, "City cannot exceed 50 characters"],
    },
    state: {
      type: String,
      trim: true,
      maxlength: [50, "State cannot exceed 50 characters"],
    },
    country: {
      type: String,
      trim: true,
      default: "India",
      maxlength: [50, "Country cannot exceed 50 characters"],
    },
    pincode: {
      type: String,
      trim: true,
      match: [/^[0-9]{6}$/, "Please enter a valid 6-digit pincode"],
    },
    profession: {
      type: String,
      trim: true,
      maxlength: [50, "Profession cannot exceed 50 characters"],
    },
    gender: {
      type: String,
      required: true,
      enum: ["Male", "Female", "Other"],
    },
    bloodgroup: {
      type: String,
      enum: ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-", null],
    },
    emgname: {
      type: String,
      trim: true,
      maxlength: [100, "Emergency contact name cannot exceed 100 characters"],
    },
    emgnumber: {
      type: String,
      trim: true,
      match: [/^[0-9]{10}$/, "Please enter a valid 10-digit phone number"],
    },
    sportperson: {
      type: Boolean,
      default: false,
    },
    sportDiscipline: {
      type: String,
      required: [true, "Sport discipline is required"],
      trim: true,
      maxlength: [100, "Sport discipline cannot exceed 100 characters"],
    },
    levelOfSport: {
      type: String,
      required: [true, "Level of sport is required"],
      trim: true,
      enum: ["Beginner", "Intermediate", "Advanced", "Professional", "Elite"],
    },
    sportsTitleEarned: {
      type: String,
      trim: true,
      maxlength: [200, "Sports title cannot exceed 200 characters"],
    },
    club: {
      type: String,
      trim: true,
      maxlength: [100, "Club name cannot exceed 100 characters"],
    },
    achievements: {
      type: String,
      maxlength: [500, "Achievements cannot exceed 500 characters"],
    },
    experienceAsCoach: {
      type: String,
      required: [true, "Experience as coach is required"],
      trim: true,
      maxlength: [50, "Experience cannot exceed 50 characters"],
    },
    overallExperience: {
      type: String,
      required: [true, "Overall experience is required"],
      trim: true,
      maxlength: [50, "Experience cannot exceed 50 characters"],
    },
    lastOrganization: {
      type: String,
      trim: true,
      maxlength: [100, "Organization name cannot exceed 100 characters"],
    },
    reference: {
      type: String,
      trim: true,
      maxlength: [100, "Reference cannot exceed 100 characters"],
    },
  },
  {
    timestamps: true,
  }
);

// All Models
const Contact = mongoose.model("Contact", contactSchema);
const Member = mongoose.model("Member", memberSchema);
const Athlete = mongoose.model("Athlete", athleteSchema);
const Coach = mongoose.model("Coach", coachSchema);

// Validation middleware
const validateRequiredFields = (requiredFields) => (req, res, next) => {
  const missingFields = requiredFields.filter((field) => !req.body[field]);

  if (missingFields.length > 0) {
    return res.status(400).json({
      success: false,
      message: `Missing required fields: ${missingFields.join(", ")}`,
      missingFields,
    });
  }
  next();
};

// All API Sections
app.get("/", (req, res) => {
  res.json({
    success: true,
    message: "Welcome to Sports Management API",
    version: "1.0.0",
    environment: NODE_ENV,
    timestamp: new Date().toISOString(),
    documentation: "/health for server status",
  });
});

// Contact API
app.post(
  "/contact",
  validateRequiredFields(["name", "email", "number", "message"]),
  async (req, res) => {
    try {
      const { name, email, number, message } = req.body;

      const newMessage = new Contact({ name, email, number, message });
      await newMessage.save();

      return res.status(201).json({
        success: true,
        message: "Contact message saved successfully",
        data: {
          id: newMessage._id,
          name: newMessage.name,
          email: newMessage.email,
          createdAt: newMessage.createdAt,
        },
      });
    } catch (error) {
      console.error("Error saving contact message:", error);

      if (error.name === "ValidationError") {
        const errors = Object.values(error.errors).map((err) => err.message);
        return res.status(400).json({
          success: false,
          message: "Validation failed",
          errors,
        });
      }

      return res.status(500).json({
        success: false,
        message: "Failed to save contact message",
        error: NODE_ENV === "development" ? error.message : undefined,
      });
    }
  }
);

app.get("/contact", async (req, res) => {
  try {
    const { page = 1, limit = 10, search } = req.query;
    const pageNum = parseInt(page);
    const limitNum = parseInt(limit);
    const skip = (pageNum - 1) * limitNum;

    // Build filter for search
    const filter = {};
    if (search) {
      filter.$or = [
        { name: { $regex: search, $options: "i" } },
        { email: { $regex: search, $options: "i" } },
        { number: { $regex: search, $options: "i" } },
        { message: { $regex: search, $options: "i" } },
      ];
    }

    const [data, total] = await Promise.all([
      Contact.find(filter)
        .sort({ createdAt: -1 })
        .skip(skip)
        .limit(limitNum)
        .lean(),
      Contact.countDocuments(filter),
    ]);

    return res.status(200).json({
      success: true,
      data,
      pagination: {
        page: pageNum,
        limit: limitNum,
        total,
        pages: Math.ceil(total / limitNum),
      },
    });
  } catch (error) {
    console.error("Error fetching contacts:", error);
    return res.status(500).json({
      success: false,
      message: "Failed to fetch contacts",
      error: NODE_ENV === "development" ? error.message : undefined,
    });
  }
});

// Member API
app.post("/member", async (req, res) => {
  try {
    const {
      aadharnumber,
      email,
      salutation,
      name,
      dob,
      number,
      address,
      city,
      state,
      country,
      pincode,
      profession,
      gender,
      bloodgroup,
      emgname,
      emgnumber,
      sportperson,
    } = req.body;

    // Validate required fields
    if (!aadharnumber || !name || !email || !number) {
      return res.status(400).json({
        success: false,
        message: "Missing required fields: aadharnumber, name, email, number",
        requiredFields: ["aadharnumber", "name", "email", "number"],
      });
    }

    // Check if member already exists
    const existingMember = await Member.findOne({
      $or: [
        { aadharnumber: aadharnumber.trim() },
        { email: email.toLowerCase().trim() },
      ],
    });

    if (existingMember) {
      return res.status(409).json({
        success: false,
        message: "Member already exists with this Aadhar number or email",
        conflict:
          existingMember.aadharnumber === aadharnumber.trim()
            ? "aadharnumber"
            : "email",
      });
    }

    // Create new member
    const newMember = new Member({
      aadharnumber: aadharnumber.trim(),
      email: email.toLowerCase().trim(),
      salutation: salutation?.trim() || "Mr",
      name: name.trim(),
      dob: dob || null,
      number: number.trim(),
      address: address?.trim(),
      city: city?.trim(),
      state: state?.trim(),
      country: country?.trim() || "India",
      pincode: pincode?.trim(),
      profession: profession?.trim(),
      gender: gender?.trim(),
      bloodgroup: bloodgroup?.trim().toUpperCase(),
      emgname: emgname?.trim(),
      emgnumber: emgnumber?.trim(),
      sportperson: sportperson || false,
    });

    // Save to database
    const savedMember = await newMember.save();

    return res.status(201).json({
      success: true,
      message: "Member created successfully",
      data: {
        id: savedMember._id,
        aadharnumber: savedMember.aadharnumber,
        name: savedMember.name,
        email: savedMember.email,
        number: savedMember.number,
        createdAt: savedMember.createdAt,
      },
    });
  } catch (error) {
    console.error("Error creating member:", error);

    if (error.name === "ValidationError") {
      const errors = Object.values(error.errors).map((err) => err.message);
      return res.status(400).json({
        success: false,
        message: "Validation failed",
        errors: errors,
      });
    }

    if (error.code === 11000) {
      return res.status(409).json({
        success: false,
        message: "Member already exists with this Aadhar number or email",
      });
    }

    return res.status(500).json({
      success: false,
      message: "Internal server error",
      error:
        NODE_ENV === "development" ? error.message : "Something went wrong",
    });
  }
});

app.get("/member", async (req, res) => {
  try {
    const {
      page = 1,
      limit = 10,
      search,
      sortBy = "createdAt",
      sortOrder = "desc",
      gender,
      bloodgroup,
      profession,
      city,
      sportperson,
    } = req.query;

    // Build filter object
    const filter = {};

    // Search across multiple fields
    if (search) {
      filter.$or = [
        { name: { $regex: search, $options: "i" } },
        { email: { $regex: search, $options: "i" } },
        { aadharnumber: { $regex: search, $options: "i" } },
        { number: { $regex: search, $options: "i" } },
        { city: { $regex: search, $options: "i" } },
        { profession: { $regex: search, $options: "i" } },
      ];
    }

    // Add filters
    if (gender) filter.gender = gender;
    if (bloodgroup) filter.bloodgroup = bloodgroup;
    if (profession) filter.profession = { $regex: profession, $options: "i" };
    if (city) filter.city = { $regex: city, $options: "i" };
    if (sportperson !== undefined) filter.sportperson = sportperson === "true";

    // Sort configuration
    const sortConfig = {};
    sortConfig[sortBy] = sortOrder === "asc" ? 1 : -1;

    // Pagination calculation
    const pageNum = parseInt(page);
    const limitNum = parseInt(limit);
    const skip = (pageNum - 1) * limitNum;

    // Execute query with pagination
    const members = await Member.find(filter)
      .sort(sortConfig)
      .skip(skip)
      .limit(limitNum)
      .select("-__v")
      .lean();

    // Get total count for pagination
    const totalMembers = await Member.countDocuments(filter);
    const totalPages = Math.ceil(totalMembers / limitNum);

    // If no members found
    if (members.length === 0) {
      return res.status(200).json({
        success: true,
        message: "No members found",
        data: {
          members: [],
          pagination: {
            currentPage: pageNum,
            totalPages,
            totalMembers,
            hasNext: pageNum < totalPages,
            hasPrev: pageNum > 1,
          },
        },
      });
    }

    // Transform data for response
    const transformedMembers = members.map((member) => ({
      id: member._id,
      aadharnumber: member.aadharnumber,
      salutation: member.salutation,
      name: member.name,
      email: member.email,
      dob: member.dob ? member.dob.toISOString().split("T")[0] : null,
      number: member.number,
      address: member.address,
      city: member.city,
      state: member.state,
      country: member.country,
      pincode: member.pincode,
      profession: member.profession,
      gender: member.gender,
      bloodgroup: member.bloodgroup,
      emgname: member.emgname,
      emgnumber: member.emgnumber,
      sportperson: member.sportperson,
      createdAt: member.createdAt,
      updatedAt: member.updatedAt,
    }));

    return res.status(200).json({
      success: true,
      message: "Members retrieved successfully",
      data: {
        members: transformedMembers,
        pagination: {
          currentPage: pageNum,
          totalPages,
          totalMembers,
          hasNext: pageNum < totalPages,
          hasPrev: pageNum > 1,
          limit: limitNum,
        },
      },
    });
  } catch (error) {
    console.error("Error fetching members:", error);
    return res.status(500).json({
      success: false,
      message: "Internal server error",
      error:
        NODE_ENV === "development" ? error.message : "Something went wrong",
    });
  }
});

// Athlete API
app.post("/athlete", async (req, res) => {
  try {
    const {
      aadharnumber,
      email,
      salutation,
      name,
      dob,
      number,
      address,
      city,
      state,
      country,
      pincode,
      profession,
      gender,
      bloodgroup,
      emgname,
      emgnumber,
      sportDiscipline,
      levelOfSport,
      sportsTitleEarned,
      club,
      achievements,
    } = req.body;

    // Validate required fields
    const requiredFields = [
      "aadharnumber",
      "name",
      "email",
      "number",
      "sportDiscipline",
      "levelOfSport",
    ];
    const missingFields = requiredFields.filter((field) => !req.body[field]);

    if (missingFields.length > 0) {
      return res.status(400).json({
        success: false,
        message: `Missing required fields: ${missingFields.join(", ")}`,
        missingFields,
      });
    }

    // Check if athlete already exists with same aadhar or email
    const existingAthlete = await Athlete.findOne({
      $or: [
        { aadharnumber: aadharnumber.trim() },
        { email: email.toLowerCase().trim() },
      ],
    });

    if (existingAthlete) {
      const conflictField =
        existingAthlete.aadharnumber === aadharnumber.trim()
          ? "aadharnumber"
          : "email";
      return res.status(409).json({
        success: false,
        message: `Athlete already exists with this ${conflictField}`,
        conflict: conflictField,
      });
    }

    // Create new athlete
    const newAthlete = new Athlete({
      aadharnumber: aadharnumber.trim(),
      email: email.toLowerCase().trim(),
      salutation: salutation?.trim() || "Mr",
      name: name.trim(),
      dob: dob || null,
      number: number.trim(),
      address: address?.trim(),
      city: city?.trim(),
      state: state?.trim(),
      country: country?.trim() || "India",
      pincode: pincode?.trim(),
      profession: profession?.trim(),
      gender: gender?.trim(),
      bloodgroup: bloodgroup?.trim()?.toUpperCase(),
      emgname: emgname?.trim(),
      emgnumber: emgnumber?.trim(),
      sportperson: true,
      sportDiscipline: sportDiscipline.trim(),
      levelOfSport: levelOfSport.trim(),
      sportsTitleEarned: sportsTitleEarned?.trim(),
      club: club?.trim(),
      achievements: achievements,
    });

    // Save to database
    const savedAthlete = await newAthlete.save();

    return res.status(201).json({
      success: true,
      message: "Athlete registered successfully",
      data: {
        id: savedAthlete._id,
        aadharnumber: savedAthlete.aadharnumber,
        name: savedAthlete.name,
        email: savedAthlete.email,
        sportDiscipline: savedAthlete.sportDiscipline,
        levelOfSport: savedAthlete.levelOfSport,
        createdAt: savedAthlete.createdAt,
      },
    });
  } catch (error) {
    console.error("Error creating athlete:", error);

    if (error.name === "ValidationError") {
      const errors = Object.values(error.errors).map((err) => err.message);
      return res.status(400).json({
        success: false,
        message: "Validation failed",
        errors,
      });
    }

    if (error.code === 11000) {
      return res.status(409).json({
        success: false,
        message: "Athlete already exists with this Aadhar number or email",
      });
    }

    return res.status(500).json({
      success: false,
      message: "Internal server error",
      error:
        NODE_ENV === "development" ? error.message : "Something went wrong",
    });
  }
});

app.get("/athlete", async (req, res) => {
  try {
    const {
      page = 1,
      limit = 10,
      search,
      sortBy = "createdAt",
      sortOrder = "desc",
      sportDiscipline,
      levelOfSport,
      gender,
      city,
      bloodgroup,
    } = req.query;

    // Build filter object
    const filter = {};

    // Search across multiple fields
    if (search) {
      filter.$or = [
        { name: { $regex: search, $options: "i" } },
        { email: { $regex: search, $options: "i" } },
        { aadharnumber: { $regex: search, $options: "i" } },
        { number: { $regex: search, $options: "i" } },
        { city: { $regex: search, $options: "i" } },
        { profession: { $regex: search, $options: "i" } },
        { sportDiscipline: { $regex: search, $options: "i" } },
        { club: { $regex: search, $options: "i" } },
        { sportsTitleEarned: { $regex: search, $options: "i" } },
      ];
    }

    // Add filters
    if (sportDiscipline)
      filter.sportDiscipline = { $regex: sportDiscipline, $options: "i" };
    if (levelOfSport) filter.levelOfSport = levelOfSport;
    if (gender) filter.gender = gender;
    if (city) filter.city = { $regex: city, $options: "i" };
    if (bloodgroup) filter.bloodgroup = bloodgroup;

    // Sort configuration
    const sortConfig = {};
    const validSortFields = [
      "name",
      "email",
      "createdAt",
      "sportDiscipline",
      "levelOfSport",
      "city",
    ];
    const sortField = validSortFields.includes(sortBy) ? sortBy : "createdAt";
    sortConfig[sortField] = sortOrder === "asc" ? 1 : -1;

    // Pagination calculation
    const pageNum = parseInt(page);
    const limitNum = parseInt(limit);
    const skip = (pageNum - 1) * limitNum;

    // Execute query with pagination
    const athletes = await Athlete.find(filter)
      .sort(sortConfig)
      .skip(skip)
      .limit(limitNum)
      .select("-__v")
      .lean();

    // Get total count for pagination
    const totalAthletes = await Athlete.countDocuments(filter);
    const totalPages = Math.ceil(totalAthletes / limitNum);

    if (athletes.length === 0) {
      return res.status(200).json({
        success: true,
        message: "No athletes found",
        data: {
          athletes: [],
          pagination: {
            currentPage: pageNum,
            totalPages,
            totalAthletes,
            hasNext: pageNum < totalPages,
            hasPrev: pageNum > 1,
          },
        },
      });
    }

    // Transform data for response
    const transformedAthletes = athletes.map((athlete) => ({
      id: athlete._id,
      aadharnumber: athlete.aadharnumber,
      salutation: athlete.salutation,
      name: athlete.name,
      email: athlete.email,
      dob: athlete.dob ? athlete.dob.toISOString().split("T")[0] : null,
      number: athlete.number,
      address: athlete.address,
      city: athlete.city,
      state: athlete.state,
      country: athlete.country,
      pincode: athlete.pincode,
      profession: athlete.profession,
      gender: athlete.gender,
      bloodgroup: athlete.bloodgroup,
      emgname: athlete.emgname,
      emgnumber: athlete.emgnumber,
      sportperson: athlete.sportperson,
      sportDiscipline: athlete.sportDiscipline,
      levelOfSport: athlete.levelOfSport,
      sportsTitleEarned: athlete.sportsTitleEarned,
      club: athlete.club,
      achievements: athlete.achievements,
      createdAt: athlete.createdAt,
      updatedAt: athlete.updatedAt,
    }));

    return res.status(200).json({
      success: true,
      message: "Athletes retrieved successfully",
      data: {
        athletes: transformedAthletes,
        pagination: {
          currentPage: pageNum,
          totalPages,
          totalAthletes,
          hasNext: pageNum < totalPages,
          hasPrev: pageNum > 1,
          limit: limitNum,
        },
      },
    });
  } catch (error) {
    console.error("Error fetching athletes:", error);

    return res.status(500).json({
      success: false,
      message: "Internal server error",
      error:
        NODE_ENV === "development" ? error.message : "Something went wrong",
    });
  }
});

// Coach API
app.post("/coach", async (req, res) => {
  try {
    const {
      aadharnumber,
      email,
      salutation,
      name,
      dob,
      number,
      address,
      city,
      state,
      country,
      pincode,
      profession,
      gender,
      bloodgroup,
      emgname,
      emgnumber,
      sportperson,
      sportDiscipline,
      levelOfSport,
      sportsTitleEarned,
      club,
      achievements,
      experienceAsCoach,
      overallExperience,
      lastOrganization,
      reference,
    } = req.body;

    // Validate required fields
    const requiredFields = [
      "aadharnumber",
      "name",
      "email",
      "number",
      "sportDiscipline",
      "levelOfSport",
      "experienceAsCoach",
      "overallExperience",
    ];
    const missingFields = requiredFields.filter((field) => !req.body[field]);

    if (missingFields.length > 0) {
      return res.status(400).json({
        success: false,
        message: `Missing required fields: ${missingFields.join(", ")}`,
        missingFields,
      });
    }

    // Check for existing coach
    const existingCoach = await Coach.findOne({
      $or: [
        { aadharnumber: aadharnumber.trim() },
        { email: email.toLowerCase().trim() },
      ],
    });

    if (existingCoach) {
      return res.status(409).json({
        success: false,
        message: "Coach already exists with this Aadhar number or email",
      });
    }

    // Create new coach
    const newCoach = new Coach({
      aadharnumber: aadharnumber.trim(),
      email: email.toLowerCase().trim(),
      salutation: salutation?.trim() || "Mr",
      name: name.trim(),
      dob: dob || null,
      number: number.trim(),
      address: address?.trim(),
      city: city?.trim(),
      state: state?.trim(),
      country: country?.trim() || "India",
      pincode: pincode?.trim(),
      profession: profession?.trim(),
      gender: gender?.trim(),
      bloodgroup: bloodgroup?.trim()?.toUpperCase(),
      emgname: emgname?.trim(),
      emgnumber: emgnumber?.trim(),
      sportperson: sportperson || false,
      sportDiscipline: sportDiscipline.trim(),
      levelOfSport: levelOfSport.trim(),
      sportsTitleEarned: sportsTitleEarned?.trim(),
      club: club?.trim(),
      achievements: achievements,
      experienceAsCoach: experienceAsCoach.trim(),
      overallExperience: overallExperience.trim(),
      lastOrganization: lastOrganization?.trim(),
      reference: reference?.trim(),
    });

    // Save to database
    const savedCoach = await newCoach.save();

    return res.status(201).json({
      success: true,
      message: "Coach registered successfully",
      data: {
        id: savedCoach._id,
        aadharnumber: savedCoach.aadharnumber,
        name: savedCoach.name,
        email: savedCoach.email,
        sportDiscipline: savedCoach.sportDiscipline,
        experienceAsCoach: savedCoach.experienceAsCoach,
        createdAt: savedCoach.createdAt,
      },
    });
  } catch (error) {
    console.error("Error creating coach:", error);

    if (error.name === "ValidationError") {
      const errors = Object.values(error.errors).map((err) => err.message);
      return res.status(400).json({
        success: false,
        message: "Validation failed",
        errors,
      });
    }

    if (error.code === 11000) {
      return res.status(409).json({
        success: false,
        message: "Coach already exists with this Aadhar number or email",
      });
    }

    return res.status(500).json({
      success: false,
      message: "Internal server error",
      error:
        NODE_ENV === "development" ? error.message : "Something went wrong",
    });
  }
});

app.get("/coach", async (req, res) => {
  try {
    const {
      page = 1,
      limit = 10,
      search,
      sortBy = "createdAt",
      sortOrder = "desc",
      sportDiscipline,
      levelOfSport,
      gender,
      city,
      experienceAsCoach,
    } = req.query;

    // Build filter object
    const filter = {};

    // Search across multiple fields
    if (search) {
      filter.$or = [
        { name: { $regex: search, $options: "i" } },
        { email: { $regex: search, $options: "i" } },
        { aadharnumber: { $regex: search, $options: "i" } },
        { number: { $regex: search, $options: "i" } },
        { city: { $regex: search, $options: "i" } },
        { profession: { $regex: search, $options: "i" } },
        { sportDiscipline: { $regex: search, $options: "i" } },
        { club: { $regex: search, $options: "i" } },
        { lastOrganization: { $regex: search, $options: "i" } },
      ];
    }

    // Add filters
    if (sportDiscipline)
      filter.sportDiscipline = { $regex: sportDiscipline, $options: "i" };
    if (levelOfSport) filter.levelOfSport = levelOfSport;
    if (gender) filter.gender = gender;
    if (city) filter.city = { $regex: city, $options: "i" };
    if (experienceAsCoach)
      filter.experienceAsCoach = { $regex: experienceAsCoach, $options: "i" };

    // Sort configuration
    const sortConfig = {};
    const validSortFields = [
      "name",
      "email",
      "createdAt",
      "sportDiscipline",
      "levelOfSport",
      "city",
      "experienceAsCoach",
    ];
    const sortField = validSortFields.includes(sortBy) ? sortBy : "createdAt";
    sortConfig[sortField] = sortOrder === "asc" ? 1 : -1;

    // Pagination calculation
    const pageNum = parseInt(page);
    const limitNum = parseInt(limit);
    const skip = (pageNum - 1) * limitNum;

    // Execute query with pagination
    const coaches = await Coach.find(filter)
      .sort(sortConfig)
      .skip(skip)
      .limit(limitNum)
      .select("-__v")
      .lean();

    // Get total count for pagination
    const totalCoaches = await Coach.countDocuments(filter);
    const totalPages = Math.ceil(totalCoaches / limitNum);

    if (coaches.length === 0) {
      return res.status(200).json({
        success: true,
        message: "No coaches found",
        data: {
          coaches: [],
          pagination: {
            currentPage: pageNum,
            totalPages,
            totalCoaches,
            hasNext: pageNum < totalPages,
            hasPrev: pageNum > 1,
          },
        },
      });
    }

    // Transform data for response
    const transformedCoaches = coaches.map((coach) => ({
      id: coach._id,
      aadharnumber: coach.aadharnumber,
      salutation: coach.salutation,
      name: coach.name,
      email: coach.email,
      dob: coach.dob ? coach.dob.toISOString().split("T")[0] : null,
      number: coach.number,
      address: coach.address,
      city: coach.city,
      state: coach.state,
      country: coach.country,
      pincode: coach.pincode,
      profession: coach.profession,
      gender: coach.gender,
      bloodgroup: coach.bloodgroup,
      emgname: coach.emgname,
      emgnumber: coach.emgnumber,
      sportperson: coach.sportperson,
      sportDiscipline: coach.sportDiscipline,
      levelOfSport: coach.levelOfSport,
      sportsTitleEarned: coach.sportsTitleEarned,
      club: coach.club,
      achievements: coach.achievements,
      experienceAsCoach: coach.experienceAsCoach,
      overallExperience: coach.overallExperience,
      lastOrganization: coach.lastOrganization,
      reference: coach.reference,
      createdAt: coach.createdAt,
      updatedAt: coach.updatedAt,
    }));

    return res.status(200).json({
      success: true,
      message: "Coaches retrieved successfully",
      data: {
        coaches: transformedCoaches,
        pagination: {
          currentPage: pageNum,
          totalPages,
          totalCoaches,
          hasNext: pageNum < totalPages,
          hasPrev: pageNum > 1,
          limit: limitNum,
        },
      },
    });
  } catch (error) {
    console.error("Error fetching coaches:", error);

    return res.status(500).json({
      success: false,
      message: "Internal server error",
      error:
        NODE_ENV === "development" ? error.message : "Something went wrong",
    });
  }
});

// Secure Check endpoint
app.post("/secureCheck", (req, res) => {
  try {
    const { key } = req.body;

    if (!key) {
      return res.status(400).json({
        success: false,
        error: "Key is required",
      });
    }

    const match = key === secureKey;
    res.json({
      success: true,
      match,
    });
  } catch (error) {
    console.error("Secure check error:", error);
    res.status(500).json({
      success: false,
      error: "Internal server error",
    });
  }
});

// Generic Excel download function
const downloadExcel = async (Model, filename, res) => {
  try {
    const data = await Model.find().lean();

    if (!data || data.length === 0) {
      return res.status(404).json({
        success: false,
        message: `No data found for ${filename}`,
      });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(filename);

    // Get headers from first document, excluding MongoDB internal fields
    const headers = ["S.No."];
    const fieldNames = Object.keys(data[0]).filter(
      (key) => !["_id", "__v"].includes(key)
    );
    headers.push(...fieldNames);

    worksheet.addRow(headers);

    // Style header
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE6E6FA" },
    };

    // Add data rows
    data.forEach((item, index) => {
      const rowData = [index + 1];

      fieldNames.forEach((fieldName) => {
        let value = item[fieldName];

        // Handle different data types
        if (value instanceof Date) {
          value = value.toLocaleDateString("en-IN");
        } else if (value === null || value === undefined) {
          value = "";
        } else if (typeof value === "object") {
          value = JSON.stringify(value);
        } else if (fieldName === "dob" && value) {
          value = new Date(value).toLocaleDateString("en-IN");
        }

        rowData.push(value);
      });

      worksheet.addRow(rowData);
    });

    // Auto-fit columns
    worksheet.columns.forEach((column, index) => {
      if (index === 0) {
        column.width = 8; // S.No. column
      } else {
        const header = headers[index];
        column.width = Math.max(15, header.length + 5);
      }
    });

    const timestamp = new Date().toISOString().split("T")[0];
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=${filename}-${timestamp}.xlsx`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(`Error generating Excel for ${filename}:`, error);
    res.status(500).json({
      success: false,
      message: "Failed to generate Excel file",
      error: NODE_ENV === "development" ? error.message : undefined,
    });
  }
};

// Excel download routes
app.get("/download/contact", (req, res) =>
  downloadExcel(Contact, "contacts", res)
);
app.get("/download/member", (req, res) =>
  downloadExcel(Member, "members", res)
);
app.get("/download/athlete", (req, res) =>
  downloadExcel(Athlete, "athletes", res)
);
app.get("/download/coach", (req, res) => downloadExcel(Coach, "coaches", res));

// Error Handling Middleware
app.use((error, req, res, next) => {
  console.error("Unhandled Error:", error);

  if (error.type === "entity.parse.failed") {
    return res.status(400).json({
      success: false,
      message: "Invalid JSON in request body",
    });
  }

  res.status(500).json({
    success: false,
    message: "Internal Server Error",
    ...(NODE_ENV === "development" && { error: error.message }),
  });
});

// 404 Handler
app.use((req, res) => {
  res.status(404).json({
    success: false,
    message: `Route ${req.method} ${req.originalUrl} not found`
  });
});

// Graceful shutdown
process.on("SIGINT", async () => {
  console.log("🛑 Received SIGINT. Shutting down gracefully...");
  await mongoose.connection.close();
  process.exit(0);
});

process.on("SIGTERM", async () => {
  console.log("🛑 Received SIGTERM. Shutting down gracefully...");
  await mongoose.connection.close();
  process.exit(0);
});

// Start server
app.listen(port, async () => {
  console.log(`🚀 Server running on port http://localhost:${port}`);
  console.log(`🌍 Environment: ${NODE_ENV}`);

  try {
    await DBCon();
  } catch (error) {
    console.error("💥 Failed to start server:", error);
    process.exit(1);
  }
});

export default app;
